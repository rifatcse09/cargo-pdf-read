<?php

namespace App\Assistants;

use Carbon\Carbon;
use Illuminate\Support\Arr;
use Illuminate\Support\Str;

class ZieglerPdfAssistant extends PdfClient
{
    /**
     * Decide if this assistant should handle the PDF.
     * AutoPdfAssistant calls ::validateFormat($lines) on each candidate.
     *
     * @param array<int,string> $lines
     * @return bool
     */
    public static function validateFormat(array $lines)
    {
        $norm = array_map(fn($l) => Str::upper(trim((string)$l)), $lines);

        $hardMarkers = [
            'ZIEGLER UK LTD',
            'BOOKING INSTRUCTION',
            'PLEASE QUOTE OUR REFERENCE',
            'BIFA', // appears in Ziegler footer/terms
        ];

        $hits = 0;
        foreach ($norm as $l) {
            foreach ($hardMarkers as $m) {
                if (Str::contains($l, $m)) { $hits++; break; }
            }
        }
        if ($hits < 2) return false;

        $soft = 0;
        foreach ($norm as $l) {
            if (Str::contains($l, 'COLLECTION') || Str::contains($l, 'DELIVERY')) $soft++;
            if (preg_match('/\b(ORDER|REF|REFERENCE|BOOKING)\b/', $l)) $soft++;
            if (preg_match('/\b(EUR|GBP|USD|€|£|\$)\b/', $l)) $soft++;
        }
        return $soft >= 2;
    }

    /**
     * Build and return the schema-shaped order array.
     * Signature MUST match PdfClient: (array $lines, ?string $attachment_filename = null)
     *
     * @param  array<int,string> $lines
     * @param  string|null $attachment_filename
     * @return array<string,mixed>
     */
    public function processLines(array $lines, ?string $attachment_filename = null)
    {
        // --- Normalize input ---
        $lines = array_values(array_filter(array_map(
            fn($l) => trim(preg_replace('/[ \t]+/u', ' ', (string)$l)),
            $lines
        ), fn($l) => $l !== ''));

        // --- Extract fields (robust/forgiving regex & heuristics) ---
        $orderRef                         = $this->extractOrderRef($lines);
        [$freightAmount, $freightCurrency]= $this->extractFreight($lines);
        $customer                         = $this->extractCustomer($lines);
        $loadingStops                     = $this->extractStops($lines, 'collection');
        $deliveryStops                    = $this->extractStops($lines, 'delivery');
        $cargos                           = $this->extractCargo($lines);
        $comment                          = $this->extractComment($lines, $orderRef);

        // --- Build payload per storage/order_schema.json ---
        $payload = [
            'attachment_filenames'   => array_values(array_filter([$attachment_filename ?? $this->currentFilename ?? null])),
            'customer'               => ['side' => 'none', 'details' => $customer],
            'order_reference'        => $orderRef,
            'freight_price'          => $freightAmount,
            'freight_currency'       => $freightCurrency,
            'loading_locations'      => $loadingStops,
            'destination_locations'  => $deliveryStops,
            'cargos'                 => $cargos,
            'comment'                => $comment,
        ];

        // Let base class finalize/validate; it returns void.
        $this->createOrder($payload);

        // IMPORTANT: return the array because tools/linters expect an array here.
        return $payload;
    }

    // -------------------------
    // Extractors
    // -------------------------

    protected function extractOrderRef(array $lines)
    {
        foreach ($lines as $l) {
            if (preg_match('/\b(Order\s*Ref|Our\s*Ref|Reference|Booking)\b[:#\s\-]*([A-Z0-9\-\/]{4,})/i', $l, $m)) {
                return trim($m[2]);
            }
        }
        return null;
    }

    /**
     * @return array{0:?float,1:?string}
     */
    protected function extractFreight(array $lines)
    {
        $amount = null; $ccy = null;
        foreach ($lines as $l) {
            if (preg_match('/\b(EUR|USD|GBP)\b[^\d]*([\d\.,]+)\b/i', $l, $m)) {
                $ccy = Str::upper($m[1]);
                $amount = (float) str_replace(',', '', $m[2]);
                break;
            }
            if (preg_match('/([€£$])\s*([\d\.,]+)/', $l, $m)) {
                $ccy = $m[1] === '€' ? 'EUR' : ($m[1] === '£' ? 'GBP' : 'USD');
                $amount = (float) str_replace(',', '', $m[2]);
                break;
            }
        }
        return [$amount, $ccy];
    }

    protected function extractCustomer(array $lines)
    {
        $company = null; $street = null; $city = null; $post = null; $country = null; $note = null;

        foreach ($lines as $i => $l) {
            $U = Str::upper($l);
            if (!$company && (Str::contains($U, 'ZIEGLER') || preg_match('/\b(LTD|GMBH|SAS|SRL|BV|NV|SARL|INC|C\/O)\b/', $U))) {
                $company = trim($l);

                for ($j = $i + 1; $j < min($i + 6, count($lines)); $j++) {
                    $x = $lines[$j];

                    if (!$street && $this->looksLikeStreet($x)) $street = $x;

                    if ((!$city || !$post) && preg_match('/\b([A-Z \'-]+)\b[, ]+\b([A-Z0-9]{3,}\b)/i', $x, $m)) {
                        $city = trim($m[1]); $post = trim($m[2]);
                    } elseif (!$post && preg_match('/\b([A-Z0-9]{3,}\s?[A-Z0-9]{2,})\b$/', $x, $m)) {
                        $post = trim($m[1]);
                    }

                    if (!$country && preg_match('/\b(UNITED KINGDOM|ENGLAND|GREAT BRITAIN|FRANCE)\b/i', $x, $m)) {
                        $country = strtoupper($m[1][0]) === 'F' ? 'FR' : 'GB';
                    }

                    if (!$note && Str::contains(Str::lower($x), ['terms', 'please quote', 'invoice'])) {
                        $note = $x;
                    }
                }
                break;
            }
        }

        return array_filter([
            'company'        => $company,
            'street_address' => $street,
            'city'           => $city,
            'postal_code'    => $post,
            'country'        => $country,
            'comment'        => $note,
        ], fn($v) => $v !== null && $v !== '');
    }

    /**
     * @return array<int,array<string,mixed>>
     */
    protected function extractStops(array $lines, string $type)
    {
        $markers = $type === 'collection'
            ? ['collection', 'pickup', 'loading', 'load from']
            : ['delivery', 'unload', 'destination', 'ship to'];

        $idxs = [];
        foreach ($lines as $i => $l) {
            $low = Str::lower($l);
            foreach ($markers as $m) {
                if (Str::contains($low, $m)) { $idxs[] = $i; break; }
            }
        }

        $stops = [];
        foreach ($idxs as $i) {
            $cand = $this->parseStopNear($lines, $i);
            if ($cand) $stops[] = $cand;
        }

        if (!$stops) {
            for ($i = 0; $i < count($lines); $i++) {
                $cand = $this->parseStopNear($lines, $i, false);
                if ($cand) { $stops[] = $cand; $i += 3; }
            }
        }

        // De-dupe by company + window
        $uniq = [];
        foreach ($stops as $s) {
            $key = Str::lower(
                Arr::get($s, 'company_address.company', '') . '|' .
                Arr::get($s, 'company_address.postal_code', '') . '|' .
                Arr::get($s, 'time.datetime_from', '') . '|' .
                Arr::get($s, 'time.datetime_to', '')
            );
            $uniq[$key] = $s;
        }

        return array_values(array_filter($uniq));
    }

    protected function parseStopNear(array $lines, int $idx, bool $strict = true)
    {
        $company = $street = $city = $post = $country = $note = null;

        for ($i = $idx + 1; $i < min($idx + 8, count($lines)); $i++) {
            $l = trim($lines[$i]);

            if (!$company && $this->looksLikeCompany($l)) { $company = $l; continue; }
            if (!$street && $this->looksLikeStreet($l))   { $street  = $l; continue; }

            if ((!$city || !$post) && preg_match('/\b([A-Z \'-]+)\b[, ]+\b([A-Z0-9]{3,}\b)/i', $l, $m)) {
                $city = trim($m[1]); $post = trim($m[2]); continue;
            }
            if (!$post && preg_match('/\b([A-Z0-9]{3,}\s?[A-Z0-9]{2,})\b$/', $l, $m)) {
                $post = trim($m[1]); continue;
            }

            if (!$country && preg_match('/\b(UNITED KINGDOM|ENGLAND|GREAT BRITAIN|FRANCE)\b/i', $l, $m)) {
                $country = strtoupper($m[1][0]) === 'F' ? 'FR' : 'GB';
            }

            if (!$note && preg_match('/\b(REF|BOOKED FOR|ORDER)\b/i', $l)) {
                $note = $l;
            }
        }

        $time = $this->findTimeWindow($lines, $idx, $idx + 8);

        if ($strict && !$company && !$time) return null;

        $addr = array_filter([
            'company'        => $company,
            'street_address' => $street,
            'city'           => $city,
            'postal_code'    => $post,
            'country'        => $country,
            'comment'        => $note,
        ], fn($v) => $v !== null && $v !== '');

        return array_filter([
            'company_address' => $addr ?: null,
            'time'            => $time ?: null,
        ]);
    }

    protected function findTimeWindow(array $lines, int $from, int $to)
    {
        $to = min($to, count($lines) - 1);
        for ($i = $from; $i <= $to; $i++) {
            $tw = $this->parseTimeLine($lines[$i]);
            if ($tw) return $tw;
        }
        for ($i = max(0, $from - 3); $i < $from; $i++) {
            $tw = $this->parseTimeLine($lines[$i]);
            if ($tw) return $tw;
        }
        return null;
    }

    protected function parseTimeLine(string $line)
    {
        $s = Str::of($line)->replace(['–','—'], '-')->toString();

        // "27/06/2025 09:00 - 12:00" | "2025-06-27 06:00 to 12:00"
        if (preg_match('/\b(\d{1,2}[\/\.-]\d{1,2}[\/\.-]\d{2,4}|\d{4}-\d{2}-\d{2})\b\s+(\d{1,2}:\d{2})\s*[-to]+\s*(\d{1,2}:\d{2})/i', $s, $m)) {
            $date = $this->parseDate($m[1]);
            return ['datetime_from' => $this->fmtDt($date, $m[2]), 'datetime_to' => $this->fmtDt($date, $m[3])];
        }
        // single date+time
        if (preg_match('/\b(\d{1,2}[\/\.-]\d{1,2}[\/\.-]\d{2,4}|\d{4}-\d{2}-\d{2})\b\s+(\d{1,2}:\d{2})\b/i', $s, $m)) {
            $date = $this->parseDate($m[1]);
            return ['datetime_from' => $this->fmtDt($date, $m[2])];
        }
        return null;
    }

    protected function parseDate(string $s): Carbon
    {
        $s = trim($s);
        if (preg_match('/^\d{4}-\d{2}-\d{2}$/', $s)) {
            return Carbon::createFromFormat('Y-m-d', $s)->startOfMinute();
        }
        $s = Str::replace(['.', '-'], '/', $s);
        [$a, $b, $c] = array_map('intval', explode('/', $s));
        if ($a > 12) {
            return Carbon::createFromFormat('d/m/Y', sprintf('%02d/%02d/%04d', $a, $b, $c))->startOfMinute();
        }
        return Carbon::createFromFormat('d/m/Y', sprintf('%02d/%02d/%04d', $a, $b, $c))->startOfMinute();
    }

    protected function fmtDt(Carbon $date, string $time): string
    {
        [$H, $M] = array_map('intval', explode(':', $time));
        return $date->copy()->setTime($H, $M, 0)->toIso8601String();
    }

    protected function looksLikeCompany(string $line): bool
    {
        $u = Str::upper($line);
        return Str::contains($u, ['LTD','INC','SAS','GMBH','SRL','BV','NV','SARL','S.A.','C/O'])
            || (preg_match('/[A-Z]{3,}/', $u) && Str::length($u) <= 80);
    }

    protected function looksLikeStreet(string $line): bool
    {
        return (bool) preg_match('/\b(ROAD|RD|ST|STREET|LANE|AVE|AVENUE|RUE|CHEMIN|CHEM\.|WAY|PARK|COURT|RTE|DR|PLACE|BLVD|GATEWAY|CROSSING)\b/i', $line);
    }

    /**
     * @return array<int,array<string,mixed>>
     */
    protected function extractCargo(array $lines)
    {
        foreach ($lines as $l) {
            if (preg_match('/\b(pallets?)\b[: ]+(\d{1,4})\b/i', $l, $m) || preg_match('/\b(\d{1,4})\s+(pallets?)\b/i', $l, $m)) {
                $count = (int) ($m[2] ?? $m[1]);
                return [[ 'title' => 'Palletized goods', 'package_count' => $count, 'package_type' => 'pallet' ]];
            }
        }
        foreach ($lines as $l) {
            if (preg_match('/\b(\d{1,5})\s+(cartons?|boxes?|packages?)\b/i', $l, $m)) {
                return [[ 'title' => 'Packaged goods', 'package_count' => (int)$m[1], 'package_type' => 'package' ]];
            }
        }
        return [];
    }

    protected function extractComment(array $lines, ?string $orderRef)
    {
        $bits = [];
        foreach ($lines as $l) {
            $ll = Str::lower($l);
            if (Str::contains($ll, ['terms', 'permission', 'signed pod', 'please quote our reference', 'bifa'])) {
                $bits[] = $l;
            }
        }
        if ($orderRef) array_unshift($bits, "Order Ref: {$orderRef}");
        return $bits ? implode(' | ', array_unique($bits)) : null;
    }
}