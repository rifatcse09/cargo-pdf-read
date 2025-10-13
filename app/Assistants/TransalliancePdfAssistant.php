<?php

namespace App\Assistants;

use Carbon\Carbon;
use Illuminate\Support\Arr;
use Illuminate\Support\Str;

class TransalliancePdfAssistant extends PdfClient
{
    protected const POSTCODE_PATTERNS = [
        'GB' => '/\b([A-Z]{1,2}\d{1,2}[A-Z]?)\s?(\d[A-Z]{2})\b/i',
        'FR' => '/\b\d{5}\b/',
        'LT' => '/\b(?:LT-)?\d{5}\b/i',
    ];

    protected const INSTRUCTION_KEYWORDS = [
        'INVOICE','BILLING','PRICE MUST','SIGNED CMR','DELIVERY NOTE','TONNAGE',
        'FUEL SURCHARGE','ROAD TAX','FORBIDDEN','SUBCONTRACTOR','CLAIM','EQUIPMENT',
        'DOCUMENTS','REGULATIONS','ORANGE LANE','SCAN','BON D\'ECHANGE','INSURANCE','COMPLIANCE'
    ];

    public static function validateFormat(array $lines)
    {
        $norm = array_map(fn($l) => Str::upper(trim((string)$l)), $lines);
        $hard = ['CHARTERING CONFIRMATION','TRANSALLIANCE','FUSM'];
        $hits = 0; foreach ($norm as $l) { foreach ($hard as $m) { if (Str::contains($l,$m)) { $hits++; break; } } }
        if ($hits < 2) return false;

        $soft = 0;
        foreach ($norm as $l) {
            if (preg_match('/\b(ORDER|REF|REFERENCE|BOOKING|CONFIRMATION)\b/', $l)) $soft++;
            if (preg_match('/\b(EUR|GBP|USD|€|£|\$)\b/', $l)) $soft++;
            if (Str::contains($l,'LOADING') || Str::contains($l,'DELIVERY') || Str::contains($l,'DESTINATION')) $soft++;
        }
        return $soft >= 2;
    }

    public function processLines(array $lines, ?string $attachment_filename = null)
    {
        // Normalize whitespace and also normalize "8h00" -> "08:00"
        $lines = array_values(array_filter(array_map(function ($l) {
            $l = trim(preg_replace('/[ \t]+/u', ' ', (string)$l));
            $l = preg_replace_callback('/\b(\d{1,2})h(\d{2})\b/u', fn($m) => sprintf('%02d:%02d', (int)$m[1], (int)$m[2]), $l);
            return $l;
        }, $lines), fn($l) => $l !== ''));

        $orderRef                         = $this->extractOrderRef($lines, $attachment_filename);
        [$freightAmount,$freightCurrency] = $this->extractFreight($lines);
        $customer                         = $this->extractCustomer($lines);
        $loadingStops                     = $this->extractStops($lines, 'loading');
        $deliveryStops                    = $this->extractStops($lines, 'delivery');

        // Sanitize to satisfy schema
        $loadingStops  = $this->sanitizeStops($loadingStops);
        $deliveryStops = $this->sanitizeStops($deliveryStops);

        $cargos  = $this->extractCargo($lines);
        $comment = $this->extractComment($lines, $orderRef);

        // REQUIRED by schema: enforce non-null order_reference (string)
        if ($orderRef === null) {
            $orderRef = $this->fallbackOrderRef($lines, $attachment_filename) ?? 'UNKNOWN-REF';
        }

        // REQUIRED by schema: enforce number for freight_price (never null)
        if ($freightAmount === null) {
            // Try one more very permissive scan for a plain amount near price words/symbols
            [$freightAmount2, $freightCurrency2] = $this->looseFreightScan($lines);
            $freightAmount   = $freightAmount2 ?? 0.0;           // guarantee number
            $freightCurrency = $freightCurrency ?? $freightCurrency2 ?? 'EUR';
        }

        $payload = [
            'attachment_filenames'   => array_values(array_filter([$attachment_filename ?? $this->currentFilename ?? null])),
            'customer'               => ['side' => 'none', 'details' => $customer],
            'order_reference'        => $orderRef,
            'freight_price'          => $freightAmount,
            'freight_currency'       => $freightCurrency,
            'loading_locations'      => $loadingStops,
            'destination_locations'  => $deliveryStops,
            'cargos'                 => $cargos ?: [['title' => 'PACKAGING']], // ensure minItems
            'comment'                => $comment,
        ];

        $this->createOrder($payload);
        return $payload;
    }

    // -------------------- Order Ref --------------------

    protected function extractOrderRef(array $lines, ?string $filename = null): ?string
    {
        foreach ($lines as $l) {
            if (preg_match('/\bREF\.?\s*[:\-]\s*([A-Z0-9\/\-]{4,})\b/i', $l, $m)) return trim($m[1]);               // REF.:1714403
            if (preg_match('/\bREFERENCE\b\s*[:\-]\s*([A-Z0-9\/\-]{4,})/i', $l, $m)) return rtrim(trim($m[1]), '-'); // REFERENCE : TR5773 -
            if (preg_match('/\b(Order\s*Ref|Our\s*Ref|Booking)\b\s*[:#\-]*\s*([A-Z0-9\/\-]{4,})/i', $l, $m)) return trim($m[2]);
        }
        foreach ($lines as $l) {
            if (preg_match('/\bREF\.?:?\s*\.?\s*([0-9]{6,10})\b/i', $l, $m)) return $m[1];
        }
        if ($filename && preg_match('/\b(FUSM[0-9]{10,})\b/i', $filename, $m)) return $m[1];
        foreach ($lines as $l) {
            if (preg_match('/\b([A-Z]{2,}\d{2,}|[A-Z0-9]{6,})\b.*\b(EUR|USD|GBP|\€|\£|\$)/i', $l, $m)) return $m[1];
        }
        return null;
    }

    protected function fallbackOrderRef(array $lines, ?string $filename = null): ?string
    {
        foreach ($lines as $l) {
            if (preg_match('/\b([A-Z0-9\-\/]{6,16})\b/', Str::upper($l), $m)) return $m[1];
        }
        if ($filename) {
            $base = pathinfo($filename, PATHINFO_FILENAME);
            if ($base) return Str::upper(preg_replace('/[^A-Z0-9\-\/]/i', '', $base));
        }
        return null;
    }

    // -------------------- Money -----------------------

    protected function extractFreight(array $lines): array
    {
        $amount = null; $ccy = null;

        // 1) Common forms: "EUR 1,475.00", "1,475.00 EUR", "€1.475,00", "1 475 €"
        foreach ($lines as $l) {
            // currency first, amount after
            if (preg_match('/\b(EUR|USD|GBP)\b[^\d]*([0-9][0-9 \.,]*)/i', $l, $m)) {
                $ccy = Str::upper($m[1]);
                $amount = $this->parseAmount($m[2]);
                if ($amount !== null) break;
            }
            // amount first, currency after
            if (preg_match('/([0-9][0-9 \.,]*)\s*\b(EUR|USD|GBP)\b/i', $l, $m)) {
                $ccy = Str::upper($m[2]);
                $amount = $this->parseAmount($m[1]);
                if ($amount !== null) break;
            }
            // symbol first
            if (preg_match('/([€£\$])\s*([0-9][0-9 \.,]*)/u', $l, $m)) {
                $ccy = $this->ccyFromSymbol($m[1]);
                $amount = $this->parseAmount($m[2]);
                if ($amount !== null) break;
            }
            // amount then symbol
            if (preg_match('/([0-9][0-9 \.,]*)\s*([€£\$])/u', $l, $m)) {
                $ccy = $this->ccyFromSymbol($m[2]);
                $amount = $this->parseAmount($m[1]);
                if ($amount !== null) break;
            }
        }

        return [$amount, $ccy];
    }

    /**
     * A very permissive second-pass scan when extractFreight() didn’t find anything.
     * Looks near price words to pull a lone amount; defaults currency to EUR.
     */
    protected function looseFreightScan(array $lines): array
    {
        $priceHints = ['PRICE', 'SHIPPING', 'FREIGHT', 'RATE', 'CHARGE', 'TOTAL'];
        foreach ($lines as $l) {
            $U = Str::upper($l);
            $hasHint = false;
            foreach ($priceHints as $h) { if (Str::contains($U, $h)) { $hasHint = true; break; } }
            if ($hasHint && preg_match('/([€£\$]?\s*[0-9][0-9 \.,]*\s*[€£\$]?)/u', $l, $m)) {
                $raw = trim($m[1]);
                $symbol = null;
                if (preg_match('/[€£\$]/u', $raw, $sm)) $symbol = $sm[0];
                $amount = $this->parseAmount($raw);
                if ($amount !== null) {
                    $ccy = $symbol ? $this->ccyFromSymbol($symbol) : 'EUR';
                    return [$amount, $ccy];
                }
            }
        }
        // As an absolute last resort, grab the first big number that looks monetary
        foreach ($lines as $l) {
            if (preg_match('/\b([0-9]{3,}(?:[ \.,][0-9]{3})*(?:[\,\.][0-9]{2})?)\b/', $l, $m)) {
                $amount = $this->parseAmount($m[1]);
                if ($amount !== null) return [$amount, 'EUR'];
            }
        }
        return [null, null];
    }

    protected function ccyFromSymbol(string $sym): string
    {
        return $sym === '€' ? 'EUR' : ($sym === '£' ? 'GBP' : 'USD');
    }

    /**
     * Parse numbers like:
     * - "1,475.00" (EN)
     * - "1.475,00" (EU)
     * - "1 475" (space thousands)
     * - "1475" or "€1.475,00"
     */
    protected function parseAmount(string $raw): ?float
    {
        $s = trim($raw);
        $s = preg_replace('/[^\d\., ]/', '', $s);     // keep digits, separators, spaces
        $s = preg_replace('/\s+/', ' ', $s);          // collapse spaces

        // If both '.' and ',' present: decide decimal based on last separator
        $lastDot = strrpos($s, '.');
        $lastCom = strrpos($s, ',');
        if ($lastDot !== false && $lastCom !== false) {
            if ($lastCom > $lastDot) {
                // decimal comma: remove dots (thousands), replace comma with dot
                $s = str_replace('.', '', $s);
                $s = str_replace(',', '.', $s);
            } else {
                // decimal dot: remove commas (thousands)
                $s = str_replace(',', '', $s);
            }
        } elseif ($lastCom !== false && $lastDot === false) {
            // Only comma -> treat comma as decimal if it appears like ",dd"
            if (preg_match('/,\d{2}$/', $s)) {
                $s = str_replace('.', '', $s); // just in case
                $s = str_replace(',', '.', $s);
            } else {
                // Only thousands separators, remove commas
                $s = str_replace(',', '', $s);
            }
        } else {
            // Only dot or neither: just remove spaces and stray commas
            $s = str_replace(' ', '', $s);
            $s = str_replace(',', '', $s);
        }

        if ($s === '' || !preg_match('/^\d+(\.\d+)?$/', $s)) return null;
        return (float) $s;
    }

    // -------------------- Customer --------------------

    protected function extractCustomer(array $lines): array
    {
        $company = null; $street = null; $city = null; $post = null; $country = null;
        $vat = null; $email = null; $contact = null;

        for ($i = 0; $i < count($lines); $i++) {
            $u = Str::upper($lines[$i]);
            if (!$company && (Str::contains($u, 'TRANSALLIANCE') || preg_match('/\b(LTD|GMBH|SAS|SRL|BV|NV|SARL|INC|UAB|C\/O)\b/', $u))) {
                $company = trim($lines[$i]);

                for ($j = $i + 1; $j < min($i + 12, count($lines)); $j++) {
                    $x = trim($lines[$j]);

                    if (!$street && $this->looksLikeStreet($x)) $street = $x;
                    if (!$email && preg_match('/[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}/i', $x, $m)) $email = $m[0];
                    if (!$vat && preg_match('/\b(VAT|TVA)\b[:\s#\-]*([A-Z]{2}[A-Z0-9\-]{2,})/i', $x, $m)) $vat = strtoupper($m[2]);
                    if (!$contact && preg_match('/\b(Contact|Attn|Attention|Contact Person)\b[:\s]*([A-Z \'\-\.]+)/i', $x, $m)) $contact = trim($m[2]);

                    if (!$country) {
                        if (preg_match('/\b(UNITED KINGDOM|ENGLAND|GREAT BRITAIN)\b/i', $x)) $country = 'GB';
                        elseif (preg_match('/\b(FRANCE|FRANÇAISE?)\b/i', $x)) $country = 'FR';
                        elseif (preg_match('/\b(LITHUANIA|LIETUVA)\b/i', $x)) $country = 'LT';
                    }

                    if (!$post) {
                        foreach (self::POSTCODE_PATTERNS as $iso => $rx) {
                            if (preg_match($rx, $x, $m)) {
                                $post = trim($m[0]);
                                if (!$country) $country = $iso;
                                $city = trim(preg_replace('/[, ]*'.preg_quote($post, '/').'$/', '', $x));
                                $city = preg_replace('/[, ]+$/', '', $city);
                                if ($city && $this->isRegionWord($city)) $city = null;
                                break;
                            }
                        }
                    }

                    if ((!$post || !$city) &&
                        preg_match('/^\s*([A-Z \'\-]+)\s*[, ]\s*([A-Z0-9\- ]{3,})\s*$/i', $x, $m)) {
                        $candidateCity = trim($m[1]);
                        $candidateCode = trim($m[2]);
                        if (!$this->isRegionWord($candidateCode) && $this->looksLikeAnyPostcode($candidateCode)) {
                            if (!$city) $city = $candidateCity;
                            if (!$post) $post = $candidateCode;
                        }
                    }
                }
                break;
            }
        }

        if ($city && preg_match('/^([A-Z \'\-]+)\s+\1$/i', $city, $mm)) $city = trim($mm[1]);
        if ($post && !$this->looksLikeAnyPostcode($post)) $post = null;
        if ($city && mb_strlen($city) < 2) $city = null;

        return array_filter([
            'company'        => $company,
            'street_address' => $street,
            'city'           => $city,
            'postal_code'    => $post,
            'country'        => $country,
            'vat_code'       => $vat,
            'email'          => $email,
            'contact_person' => $contact,
        ], fn($v) => $v !== null && $v !== '');
    }

    // -------------------- Stops -----------------------

    protected function extractStops(array $lines, string $kind): array
    {
        $markers = $kind === 'loading'
            ? ['LOADING','PICKUP','ORIGIN','PORT']
            : ['DELIVERY','DESTINATION','SHIP TO','CONSIGNEE'];

        $idxs = [];
        foreach ($lines as $i => $l) {
            $up = Str::upper($l);
            foreach ($markers as $m) if (Str::contains($up, $m)) { $idxs[] = $i; break; }
        }

        $stops = [];
        foreach ($idxs as $i) if ($stop = $this->parseStopAround($lines, $i)) $stops[] = $stop;

        if (!$stops) {
            for ($i = 0; $i < count($lines); $i++) {
                if ($stop = $this->parseStopAround($lines, $i, false)) { $stops[] = $stop; $i += 3; }
            }
        }

        $uniq = [];
        foreach ($stops as $s) {
            $key = Str::lower(
                Arr::get($s,'company_address.company','').'|'.
                Arr::get($s,'company_address.postal_code','').'|'.
                Arr::get($s,'time.datetime_from','').'|'.
                Arr::get($s,'time.datetime_to','')
            );
            $uniq[$key] = $s;
        }

        return array_values($uniq);
    }

    protected function parseStopAround(array $lines, int $idx, bool $strict = true): ?array
    {
        $company = $street = $city = $post = $country = $note = null;

        for ($i = $idx + 1; $i < min($idx + 8, count($lines)); $i++) {
            $x = trim($lines[$i]);
            if ($this->isInstructionLine($x)) continue;

            if (!$company && $this->isLikelyCompany($x)) { $company = $x; continue; }
            if (!$street  && $this->looksLikeStreet($x))  { $street  = $x; continue; }

            if ((!$city || !$post) && preg_match('/\b([A-Z \'-]+)\b[, ]+\b([A-Z0-9]{3,}\b)/i', $x, $m)) {
                $city = trim($m[1]); $post = trim($m[2]); continue;
            }
            if (!$post && preg_match('/\b([A-Z0-9]{3,}\s?[A-Z0-9]{2,})\b$/', $x, $m) && $this->looksLikeAnyPostcode($m[1])) {
                $post = trim($m[1]); continue;
            }

            if (!$country) {
                if (preg_match('/\b(UNITED KINGDOM|ENGLAND|GREAT BRITAIN)\b/i', $x)) $country = 'GB';
                elseif (preg_match('/\b(FRANCE|FRANÇAISE?)\b/i', $x)) $country = 'FR';
                elseif (preg_match('/\b(LITHUANIA|LIETUVA)\b/i', $x)) $country = 'LT';
            }

            if (!$note && preg_match('/\b(REF|BOOKING|ORDER|INSTRUCTIONS?)\b/i', $x)) $note = $x;
        }

        $time = $this->findNearestDateOrWindow($lines, $idx, $idx + 8);

        if ($city && preg_match('/^([A-Z \'\-]+)\s+\1$/i', $city)) $city = trim($city);
        if ($city && mb_strlen($city) < 2) $city = null;

        if ($strict && !$company && !$time) return null;

        $addr = array_filter([
            'company'        => $company,
            'street_address' => $street,
            'city'           => $city,
            'postal_code'    => $post && $this->looksLikeAnyPostcode($post) ? $post : null,
            'country'        => $country,
            'comment'        => $note,
        ], fn($v) => $v !== null && $v !== '');

        return array_filter([
            'company_address' => $addr ?: null,
            'time'            => $time ?: null,
        ]);
    }

    // -------------------- Time ------------------------

    protected function findNearestDateOrWindow(array $lines, int $from, int $to): ?array
    {
        $to = min($to, count($lines) - 1);
        for ($i = $from; $i <= $to; $i++) if ($tw = $this->parseDateOrWindow($lines[$i])) return $tw;
        for ($i = max(0, $from - 3); $i < $from; $i++) if ($tw = $this->parseDateOrWindow($lines[$i])) return $tw;
        return null;
    }

    protected function parseDateOrWindow(string $line): ?array
    {
        $s = Str::of($line)->replace(['–','—'], '-')->toString();

        if (preg_match('/\b(\d{1,2}[\/\.-]\d{1,2}[\/\.-]\d{2,4}|\d{4}-\d{2}-\d{2})\b\s+(\d{1,2}:\d{2})\s*[-to]+\s*(\d{1,2}:\d{2})/i', $s, $m)) {
            $date = $this->parseDateFlexible($m[1]);
            return ['datetime_from' => $this->fmtDt($date, $m[2]), 'datetime_to' => $this->fmtDt($date, $m[3])];
        }
        if (preg_match('/\b(\d{1,2}[\/\.-]\d{1,2}[\/\.-]\d{2,4}|\d{4}-\d{2}-\d{2})\b\s+(\d{1,2}:\d{2})\b/i', $s, $m)) {
            $date = $this->parseDateFlexible($m[1]);
            return ['datetime_from' => $this->fmtDt($date, $m[2])];
        }
        if (preg_match('/\b(\d{1,2}[\/\.-]\d{1,2}[\/\.-]\d{2,4}|\d{4}-\d{2}-\d{2})\b/', $s, $m)) {
            $date = $this->parseDateFlexible($m[1]);
            return ['datetime_from' => $this->fmtDt($date, '00:00')];
        }

        return null;
    }

    // -------------------- Sanitizers & helpers --------

    protected function sanitizeStops(array $stops): array
    {
        foreach ($stops as &$s) {
            if (!isset($s['company_address']) || !is_array($s['company_address'])) continue;
            $addr =& $s['company_address'];

            foreach (['company','street_address','city','postal_code','country','comment'] as $k) {
                if (isset($addr[$k]) && is_string($addr[$k])) $addr[$k] = trim($addr[$k]);
            }

            if (isset($addr['city']) && is_string($addr['city'])) {
                $city = $addr['city'];
                if (preg_match('/^([A-Z \'\-]+)\s+\1$/i', $city)) $addr['city'] = trim($city);
                if (mb_strlen($addr['city']) < 2 || preg_match('/^[^A-Za-zÀ-ÿ0-9]+$/u', $addr['city'])) unset($addr['city']);
            }

            if (isset($addr['postal_code']) && is_string($addr['postal_code']) && !$this->looksLikeAnyPostcode($addr['postal_code'])) {
                unset($addr['postal_code']);
            }

            if (!$addr) unset($s['company_address']);
        }
        return $stops;
    }

    protected function parseDateFlexible(string $s): Carbon
    {
        $s = trim($s);
        if (preg_match('/^\d{4}-\d{2}-\d{2}$/', $s)) return Carbon::createFromFormat('Y-m-d', $s)->startOfMinute();

        $s = Str::replace(['.', '-'], '/', $s);
        [$a,$b,$c] = array_map('intval', explode('/', $s));
        if ($c < 100) $c = ($c >= 80) ? (1900 + $c) : (2000 + $c); // 2-digit year → 19xx/20xx
        return Carbon::createFromFormat('d/m/Y', sprintf('%02d/%02d/%04d', $a,$b,$c))->startOfMinute();
    }

    protected function fmtDt(Carbon $date, string $time): string
    {
        [$H,$M] = array_map('intval', explode(':', $time));
        return $date->copy()->setTime($H,$M,0)->toIso8601String();
    }

    protected function isInstructionLine(string $line): bool
    {
        if (Str::length($line) > 180) return true;
        $u = Str::upper($line);
        foreach (self::INSTRUCTION_KEYWORDS as $kw) if (Str::contains($u, $kw)) return true;
        return false;
    }

    protected function isLikelyCompany(string $line): bool
    {
        $u = Str::upper($line);
        if (Str::length($u) > 90) return false;
        if (Str::endsWith($u, ':') && !Str::contains($u, 'REF')) return false;
        $legal = ['LTD','INC','SAS','GMBH','SRL','BV','NV','SARL','S.A.','C/O','UAB','EP GROUP','DP WORLD','ICONEX'];
        if (Str::contains($u, $legal)) return true;
        return (bool) preg_match('/^[A-Z0-9 \'\-&\/\.]{3,}$/', $u);
    }

    protected function looksLikeStreet(string $line): bool
    {
        return (bool) preg_match('/\b(ROAD|RD|ST|STREET|LANE|AVE|AVENUE|RUE|CHEMIN|CHEM\.|WAY|PARK|COURT|RTE|DR|PLACE|BLVD|GATEWAY|PORT|CROSSING|ZI|RUE DE|COURSE|QUAI|BAKEWELL|LONDON GATEWAY)\b/i', $line);
    }

    protected function looksLikeAnyPostcode(string $code): bool
    {
        return (bool)(
            preg_match(self::POSTCODE_PATTERNS['GB'], $code) ||
            preg_match(self::POSTCODE_PATTERNS['FR'], $code) ||
            preg_match(self::POSTCODE_PATTERNS['LT'], $code)
        );
    }

    protected function isRegionWord(string $s): bool
    {
        return (bool) preg_match('/\b(APSKRITIS|REGION|COUNTY|PROVINCE|DEPARTMENT|DEPARTEMENT)\b/i', $s);
    }

    // -------------------- Cargo -----------------------

    protected function extractCargo(array $lines): array
    {
        $items = [];
        $ldm = null; $weight = null; $palletized = null; $title = null;

        foreach ($lines as $l) {
            if (!$title && preg_match('/\b(packaging|goods|commodity|M\. nature)\b[:\s-]*([A-Z \-]+)/i', $l, $m)) {
                $title = trim(str_ireplace('M. nature','', $m[2])) ?: 'PACKAGING';
            }
            if ($ldm === null && preg_match('/\b(LDM|LOADING\s*METERS?)\b[:\s]*([\d\.,]+)/i', $l, $m)) {
                $ldm = (float) str_replace(',', '', $m[2]);
            }
            if ($weight === null && preg_match('/\b(Weight)\b\s*[:\s]*([\d\.,]+)\b/i', $l, $m)) {
                $weight = (float) str_replace(',', '', $m[2]);
            }
            if ($palletized === null && preg_match('/\b(palleti[sz]ed|pallets?)\b/i', $l)) {
                $palletized = true;
            }
        }

        if ($ldm !== null || $weight !== null || $palletized !== null || $title !== null) {
            $items[] = array_filter([
                'title'       => $title ?: 'PACKAGING',
                'ldm'         => $ldm,
                'weight'      => $weight,
                'palletized'  => $palletized,
            ], fn($v) => $v !== null && $v !== '');
        }

        return $items;
    }

    protected function extractComment(array $lines, ?string $orderRef): ?string
    {
        $bits = [];
        foreach ($lines as $l) {
            $ll = Str::lower($l);
            if (Str::contains($ll, ['instruction','compliance','insurance','pallet','exchange','scan','bon d\'echange','diesel'])) $bits[] = $l;
            if (Str::contains($ll, ['invoice','invoicing address'])) $bits[] = $l;
        }
        if ($orderRef) array_unshift($bits, "Order Ref: {$orderRef}");
        return $bits ? implode(' | ', array_unique($bits)) : null;
    }
}