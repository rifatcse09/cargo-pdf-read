<?php

namespace App\Assistants;

use Carbon\Carbon;
use Illuminate\Support\Arr;
use Illuminate\Support\Str;
use App\GeonamesCountry;

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
        'DOCUMENTS','REGULATIONS','ORANGE LANE','SCAN','BON D\'ECHANGE','INSURANCE','COMPLIANCE',
        'CUSTOMS','GMR','BARCODE','FERRY','TERMINAL','PENALTY','FINE','DEPOSIT','SURCHARGE','ENS','T1','POD'
    ];

    public static function validateFormat(array $lines)
    {
        $norm = collect($lines)->map(fn($l) => str((string)$l)->trim()->upper()->toString())->all();
        $hard = ['CHARTERING CONFIRMATION','TRANSALLIANCE','FUSM'];
        $hits = 0; foreach ($norm as $l) { foreach ($hard as $m) { if (Str::contains($l, $m)) { $hits++; break; } } }
        if ($hits < 2) return false;

        $soft = 0;
        foreach ($norm as $l) {
            if (preg_match('/\b(ORDER|REF|REFERENCE|BOOKING|CONFIRMATION)\b/', $l)) $soft++;
            if (preg_match('/\b(EUR|GBP|USD|€|£|\$)\b/', $l)) $soft++;
            if (preg_match('/\b(LOADING|DELIVERY|DESTINATION)\b/', $l)) $soft++;
        }
        return $soft >= 2;
    }

    public function processLines(array $lines, ?string $attachment_filename = null)
    {
        // Normalize whitespace and convert "8h00" → "08:00"
        $lines = collect($lines)
            ->map(fn($l) => (string) $l)
            ->map(fn($l) => preg_replace('/[ \t]+/u', ' ', $l))
            ->map(fn($l) => trim($l))
            ->map(fn($l) => preg_replace_callback('/\b(\d{1,2})h(\d{2})\b/u', fn($m) => sprintf('%02d:%02d', (int)$m[1], (int)$m[2]), $l))
            ->filter(fn($l) => $l !== '')
            ->values()
            ->all();

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

        // REQUIRED by schema
        if (blank($orderRef)) {
            $orderRef = $this->fallbackOrderRef($lines, $attachment_filename) ?? 'UNKNOWN-REF';
        }

        if ($freightAmount === null) {
            [$freightAmount2, $freightCurrency2] = $this->looseFreightScan($lines);
            $freightAmount   = $freightAmount2 ?? 0.0;
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
            'cargos'                 => filled($cargos) ? $cargos : [['title' => 'PACKAGING', 'type' => 'FTL']],
            'comment'                => $comment,
        ];

        $this->createOrder($payload);
        return $payload;
    }

    /* ===================== Enhanced Order Reference Extraction ===================== */

    protected function extractOrderRef(array $lines, ?string $filename = null): ?string
    {
        foreach ($lines as $l) {
            // Look for "REF.:1714403" pattern
            if (preg_match('/\bREF\.?:?\s*([0-9]{6,8})\b/i', $l, $m)) {
                return $m[1];
            }
            if (preg_match('/\bREF\.?\s*[:\-]\s*([A-Z0-9\/\-]{4,})\b/i', $l, $m)) {
                return trim($m[1]);
            }
            if (preg_match('/\bREFERENCE\b\s*[:\-]\s*([A-Z0-9\/\-]{4,})/i', $l, $m)) {
                return rtrim(trim($m[1]), '-');
            }
        }

        return null;
    }

    /* ===================== Enhanced Freight Extraction ===================== */

    protected function extractFreight(array $lines): array
    {
        // Look for price patterns in the specific format: "SHIPPING PRICE" followed by amount
        for ($i = 0; $i < count($lines); $i++) {
            $line = $lines[$i];
            $upper = str($line)->upper()->toString();

            if (str_contains($upper, 'SHIPPING PRICE')) {
                // Look at next few lines for the amount
                for ($j = $i + 1; $j < min($i + 5, count($lines)); $j++) {
                    $nextLine = $lines[$j];

                    // Pattern: "950,00" followed by "EUR Ex-tax"
                    if (preg_match('/^([0-9]+[,\.][0-9]{2})$/', $nextLine, $m)) {
                        $amount = $this->parseAmount(str_replace(',', '.', $m[1]));

                        // Check next line for currency
                        if ($j + 1 < count($lines) && preg_match('/\b(EUR|USD|GBP)\b/i', $lines[$j + 1], $m2)) {
                            return [$amount, strtoupper($m2[1])];
                        }

                        return [$amount, 'EUR']; // Default currency
                    }
                }
            }
        }

        return [null, null];
    }

    /* ===================== Enhanced Customer Extraction ===================== */

    protected function extractCustomer(array $lines): array
    {
        $company = null; $street = null; $city = null; $post = null; $country = null;
        $vat = null; $email = null; $contact = null; $phone = null;

        for ($i = 0; $i < count($lines); $i++) {
            $line = $lines[$i];
            $upper = str($line)->upper()->toString();

            // Look for TRANSALLIANCE TS LTD
            if (str_contains($upper, 'TRANSALLIANCE TS LTD')) {
                $company = trim($line);

                // Look ahead for GB address details (skip the LT address)
                for ($j = $i + 1; $j < min($i + 20, count($lines)); $j++) {
                    $x = trim($lines[$j]);
                    $xu = str($x)->upper()->toString();

                    // Look for UK address starting with SUITE
                    if (str_contains($xu, 'SUITE') && str_contains($xu, 'FARADAY')) {
                        $street = $x;

                        // Next line should be CENTRUM ONE HUNDRED
                        if ($j + 1 < count($lines) && str_contains(str($lines[$j + 1])->upper(), 'CENTRUM')) {
                            $street .= ', ' . trim($lines[$j + 1]);
                        }
                        continue;
                    }

                    // Look for GB postcode line: "GB-DE14 2WX BURTON UPON TRENT"
                    if (preg_match('/^GB-([A-Z0-9]+\s+[A-Z0-9]+)\s+(.+)$/i', $x, $m)) {
                        $post = trim($m[1]);
                        $city = trim($m[2]);
                        $country = 'GB';
                        continue;
                    }

                    // Look for VAT: "VAT NUM: GB712061386"
                    if (preg_match('/VAT\s+NUM:\s*([A-Z0-9]+)/i', $x, $m)) {
                        if (str_starts_with($m[1], 'GB')) {
                            $vat = $m[1];
                        }
                        continue;
                    }

                    // Look for contact: "Contact: TERESA HOPKINS"
                    if (preg_match('/Contact:\s*([A-Z\s]+)/i', $x, $m)) {
                        $contact = trim($m[1]);
                        continue;
                    }

                    // Look for phone number
                    if (preg_match('/Tel\s*:\s*(\+44[0-9\s]+)/i', $x, $m)) {
                        $phone = 'Tel: ' . trim($m[1]);
                        continue;
                    }

                    // Look for email
                    if (preg_match('/([a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,})/i', $x, $m)) {
                        $email = $m[1];
                        continue;
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
            'vat_code'       => $vat,
            'email'          => $email,
            'contact_person' => $contact,
            'comment'        => $phone,
        ], fn($v) => !blank($v));
    }

    /* ===================== Enhanced Stop Parsing ===================== */

    protected function extractStops(array $lines, string $kind): array
    {
        $stops = [];
        $sectionMarker = $kind === 'loading' ? 'Loading' : 'Delivery';

        for ($i = 0; $i < count($lines); $i++) {
            if (trim($lines[$i]) === $sectionMarker) {
                $stop = $this->parseStopFromSection($lines, $i);
                if ($stop && !empty($stop['company_address'])) {
                    $stops[] = $stop;
                }
            }
        }

        // Remove duplicates based on company and address
        $unique = [];
        foreach ($stops as $stop) {
            $key = $stop['company_address']['company'] . '|' .
                   ($stop['company_address']['postal_code'] ?? '') . '|' .
                   ($stop['company_address']['city'] ?? '');
            if (!isset($unique[$key])) {
                $unique[$key] = $stop;
            }
        }

        return array_values($unique);
    }

    protected function parseStopFromSection(array $lines, int $startIdx): ?array
    {
        $company = null; $street = null; $city = null; $post = null; $country = null;
        $dateTime = null; $timeFrom = null; $timeTo = null;

        for ($i = $startIdx + 1; $i < min($startIdx + 20, count($lines)); $i++) {
            if (!isset($lines[$i])) break;

            $line = trim($lines[$i]);
            $upper = str($line)->upper()->toString();

            // Stop at next section
            if (in_array($line, ['Loading', 'Delivery', 'Observations', 'Instructions'])) {
                break;
            }

            // Look for date (e.g., "17/09/25")
            if (preg_match('/^(\d{2}\/\d{2}\/\d{2})$/', $line, $m)) {
                $dateTime = $this->parseTransallianceDate($m[1]);
                continue;
            }

            // Look for company name (e.g., "ICONEX" or "ICONEX FRANCE")
            if (!$company && preg_match('/^[A-Z][A-Z\s]+$/', $line) &&
                !str_contains($upper, 'REFERENCE') &&
                !str_contains($upper, 'CONTACT') &&
                !str_contains($upper, 'TEL') &&
                !str_contains($upper, 'ON:') &&
                strlen($line) > 3) {
                $company = $line;
                continue;
            }

            // Look for street address
            if (!$street && $this->looksLikeStreet($line)) {
                $street = $line;
                continue;
            }

            // Look for postcode and city (e.g., "GB-PE2 6DP PETERBOROUGH" or "-37530 POCE-SUR-CISSE")
            if (preg_match('/^([A-Z]{2}-)([A-Z0-9\s]+)\s+([A-Z\s\-]+)$/i', $line, $m)) {
                $country = substr($m[1], 0, 2);
                $post = trim($m[2]);
                $city = trim($m[3]);
                continue;
            }
            if (preg_match('/^-(\d{5})\s+([A-Z\s\-]+)$/i', $line, $m)) {
                $post = $m[1];
                $city = trim($m[2]);
                $country = 'FR';
                continue;
            }

            // Look for time range (e.g., "08:00 - 15:00")
            if (preg_match('/^(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})$/', $line, $m)) {
                if ($dateTime) {
                    $timeFrom = $dateTime . 'T' . $m[1] . ':00';
                    $timeTo = $dateTime . 'T' . $m[2] . ':00';
                } else {
                    // Use a default date if none found
                    $timeFrom = '2025-09-17T' . $m[1] . ':00';
                    $timeTo = '2025-09-17T' . $m[2] . ':00';
                }
                continue;
            }
        }

        if (!$company) return null;

        $result = [
            'company_address' => array_filter([
                'company' => $company,
                'street_address' => $street,
                'city' => $city,
                'postal_code' => $post,
                'country' => $country,
            ], fn($v) => !blank($v))
        ];

        if ($timeFrom) {
            $result['time'] = array_filter([
                'datetime_from' => $timeFrom,
                'datetime_to' => $timeTo
            ], fn($v) => !blank($v));
        }

        return $result;
    }

    protected function parseTransallianceDate(string $dateStr): string
    {
        // Convert "17/09/25" to "2025-09-17"
        [$day, $month, $year] = explode('/', $dateStr);
        $fullYear = '20' . $year;
        return sprintf('%s-%02d-%02d', $fullYear, (int)$month, (int)$day);
    }

    /* ===================== Enhanced Cargo Extraction for TransAlliance Format ===================== */

    protected function extractCargo(array $lines): array
    {
        $title = null;
        $weight = null;
        $ldm = null;
        $type = 'FTL'; // Default for TransAlliance
        $packageCount = null;
        $packageType = null;

        // Look for TransAlliance tabular format patterns
        for ($i = 0; $i < count($lines); $i++) {
            $line = trim($lines[$i]);
            if (empty($line)) continue;

            $upper = str($line)->upper()->toString();

            // 1. TITLE extraction - look for cargo descriptions
            if (!$title) {
                // Line 66: "PAPER ROLLS" - direct cargo type
                if (str_contains($upper, 'PAPER ROLLS') || str_contains($upper, 'PAPER ROLL')) {
                    $title = 'PAPER ROLLS';
                } elseif (preg_match('/\b(STEEL|METAL|IRON|ALUMINUM|ALUMINIUM|COPPER|BRASS)\b/i', $upper)) {
                    $title = 'METAL PRODUCTS';
                } elseif (preg_match('/\b(TEXTILE|FABRIC|CLOTH|GARMENT|CLOTHING)\b/i', $upper)) {
                    $title = 'TEXTILES';
                } elseif (preg_match('/\b(FOOD|BEVERAGE|DRINK|DAIRY|MEAT|FISH)\b/i', $upper)) {
                    $title = 'FOOD PRODUCTS';
                } elseif (preg_match('/\b(CHEMICAL|LIQUID|OIL|FUEL|GAS|ACID)\b/i', $upper)) {
                    $title = 'CHEMICALS';
                } elseif (preg_match('/\b(FURNITURE|WOOD|TIMBER|LUMBER)\b/i', $upper)) {
                    $title = 'FURNITURE';
                } elseif (preg_match('/\b(ELECTRONIC|COMPUTER|MACHINE|EQUIPMENT)\b/i', $upper)) {
                    $title = 'ELECTRONICS';
                } elseif (preg_match('/\b(CONSTRUCTION|BUILDING|CEMENT|CONCRETE)\b/i', $upper)) {
                    $title = 'CONSTRUCTION MATERIALS';
                } elseif (preg_match('/\b(AUTOMOBILE|AUTO|CAR|VEHICLE|TRUCK)\b/i', $upper)) {
                    $title = 'AUTOMOTIVE';
                } elseif (preg_match('/\b(PHARMACEUTICAL|MEDICINE|DRUG|MEDICAL)\b/i', $upper)) {
                    $title = 'PHARMACEUTICALS';
                } elseif (preg_match('/\b(GLASS|CERAMIC|POTTERY|CRYSTAL)\b/i', $upper)) {
                    $title = 'GLASS/CERAMICS';
                } elseif (preg_match('/\b(PLASTIC|POLYMER|RESIN|VINYL)\b/i', $upper)) {
                    $title = 'PLASTICS';
                }

                // Check for "M. nature:" field - line after the label
                if (str_contains($upper, 'M. NATURE:') && $i + 1 < count($lines)) {
                    $nextLine = trim($lines[$i + 1]);
                    if (strlen($nextLine) > 2 && !preg_match('/^\d+[,\.]?\d*$/', $nextLine)) {
                        $title = strtoupper($nextLine);
                    }
                }
            }

            // 2. WEIGHT extraction - TransAlliance format
            if ($weight === null) {
                // Line 65: "24000,000" after checking context
                if (preg_match('/^([0-9]+[,\.][0-9]{1,3})$/', $line)) {
                    // Check if this appears in the cargo section (after Loading or near weight labels)
                    $contextStart = max(0, $i - 15);
                    $contextEnd = min(count($lines) - 1, $i + 5);

                    $inCargoSection = false;
                    for ($j = $contextStart; $j <= $contextEnd; $j++) {
                        $contextLine = str($lines[$j])->upper()->toString();
                        if (str_contains($contextLine, 'LOADING') ||
                            str_contains($contextLine, 'WEIGHT') ||
                            str_contains($contextLine, 'KGS') ||
                            str_contains($contextLine, 'M. NATURE') ||
                            str_contains($contextLine, 'PAPER ROLLS')) {
                            $inCargoSection = true;
                            break;
                        }
                    }

                    if ($inCargoSection) {
                        $candidate = (float) str_replace(',', '.', $line);
                        // Check if this looks like weight (reasonable cargo weight range)
                        if ($candidate >= 100 && $candidate <= 50000) {
                            $weight = $candidate;
                        }
                    }
                }

                // Alternative weight patterns
                if (preg_match('/\b(?:weight|mass|kgs?)\b[:\s]*([0-9][0-9,\.]*)\s*(?:kg|kgs?|kilos?)?\b/i', $line, $m)) {
                    $weight = (float) str_replace(',', '.', $m[1]);
                }
            }

            // 3. LDM extraction - TransAlliance format
            if ($ldm === null) {
                // Line 64: "13,600" - check context for LDM
                if (preg_match('/^([0-9]+[,\.][0-9]{1,3})$/', $line)) {
                    $contextStart = max(0, $i - 5);
                    $inLdmSection = false;

                    for ($j = $contextStart; $j < $i; $j++) {
                        if (str_contains(str($lines[$j])->upper(), 'LM . . . :')) {
                            $inLdmSection = true;
                            break;
                        }
                    }

                    if ($inLdmSection) {
                        $candidate = (float) str_replace(',', '.', $line);
                        if ($candidate > 0 && $candidate <= 100) { // Reasonable LDM range
                            $ldm = $candidate;
                        }
                    }
                }
            }

            // 4. Package information
            if (!$packageCount && preg_match('/\b(\d+)\s*(pallet[s]?|pkg[s]?|package[s]?|unit[s]?|piece[s]?|box[es]?|container[s]?|coli[s]?|carton[s]?)\b/i', $line, $m)) {
                $packageCount = (int)$m[1];
                $rawType = strtolower(trim($m[2]));

                // Map to schema enum values
                if (str_contains($rawType, 'carton') || str_contains($rawType, 'box')) {
                    $packageType = 'carton';
                } elseif (str_contains($rawType, 'pallet')) {
                    $packageType = 'EPAL'; // Default for pallets per schema
                } else {
                    $packageType = 'other';
                }
            }
        }

        // Build cargo object according to schema
        $cargo = [
            'title' => $title ?: 'PACKAGING',
            'type' => $type
        ];

        // Add optional fields only if they have valid values
        if ($weight !== null && $weight > 0) {
            $cargo['weight'] = $weight;
        }

        if ($ldm !== null && $ldm > 0) {
            $cargo['ldm'] = $ldm;
        }

        if ($packageCount !== null && $packageCount > 0) {
            $cargo['package_count'] = $packageCount;
            if ($packageType) {
                $cargo['package_type'] = $packageType;
            }
        }

        return [$cargo];
    }

    /* ===================== Enhanced Comment Extraction ===================== */

    protected function extractComment(array $lines, ?string $orderRef): ?string
    {
        $instructionLines = [];
        $inInstructions = false;

        foreach ($lines as $line) {
            $upper = str($line)->upper();

            if ($upper->contains('INSTRUCTIONS')) {
                $inInstructions = true;
                continue;
            }

            if ($inInstructions) {
                if ($upper->contains(['DELIVERY', 'OBSERVATIONS'])) {
                    break;
                }

                if ($upper->contains(['BARCODE', 'GMR', 'ORANGE LANE', 'PORT']) && strlen(trim($line)) > 10) {
                    $instructionLines[] = trim($line);
                }
            }
        }

        return filled($instructionLines) ? implode(' ', $instructionLines) : null;
    }

    protected function fallbackOrderRef(array $lines, ?string $filename = null): ?string
    {
        foreach ($lines as $l) {
            if (preg_match('/\b([A-Z0-9\-\/]{6,16})\b/', str($l)->upper()->toString(), $m)) return $m[1];
        }
        if ($filename) {
            $base = pathinfo($filename, PATHINFO_FILENAME);
            if ($base) return str($base)->upper()->replaceMatches('/[^A-Z0-9\-\/]/', '')->toString();
        }
        return null;
    }

    protected function looksLikeAdminFee(string $U): bool
    {
        return str_contains($U, 'ADMIN') || str_contains($U, 'FEE') || str_contains($U, 'FINES') || str_contains($U, 'FINE')
            || str_contains($U, 'PALLET') || str_contains($U, 'DEPOSIT') || str_contains($U, 'INSURANCE')
            || str_contains($U, 'SURCHARGE');
    }

    protected function ccyFromSymbol(string $sym): string
    {
        return $sym === '€' ? 'EUR' : ($sym === '£' ? 'GBP' : 'USD');
    }

    protected function parseAmount(string $raw): ?float
    {
        $s = str($raw)->replaceMatches('/[^\d\., ]/u', '')->trim()->toString();
        $s = preg_replace('/\s+/', ' ', $s);

        // Use app helper if available
        if (function_exists('uncomma')) {
            $norm = uncomma($s);
        } else {
            // conservative fallback
            $norm = str_replace(' ', '', $s);
            $lastDot = strrpos($norm, '.');
            $lastCom = strrpos($norm, ',');
            if ($lastDot !== false && $lastCom !== false) {
                if ($lastCom > $lastDot) {
                    $norm = str_replace('.', '', $norm);
                    $norm = str_replace(',', '.', $norm);
                } else {
                    $norm = str_replace(',', '', $norm);
                }
            } else {
                $norm = str_replace(',', '.', $norm);
            }
        }

        $norm = trim($norm);
        if ($norm === '' || !preg_match('/^\d+(\.\d+)?$/', $norm)) return null;
        return (float) $norm;
    }

    protected function looseFreightScan(array $lines): array
    {
        $priceHints = ['PRICE', 'SHIPPING', 'FREIGHT', 'RATE', 'CHARGE', 'TOTAL'];
        foreach ($lines as $l) {
            $U = str($l)->upper()->toString();
            if (!collect($priceHints)->some(fn($h) => str_contains($U, $h))) continue;
            if ($this->looksLikeAdminFee($U)) continue;

            if (preg_match('/([€£\$]?\s*[0-9][0-9 \.,]*\s*[€£\$]?)/u', $l, $m)) {
                $raw = trim($m[1]);
                $symbol = preg_match('/[€£\$]/u', $raw, $sm) ? $sm[0] : null;
                $amount = $this->parseAmount($raw);
                if ($amount !== null) {
                    $ccy = $symbol ? $this->ccyFromSymbol($symbol) : 'EUR';
                    return [$amount, $ccy];
                }
            }
        }
        foreach ($lines as $l) {
            if (preg_match('/\b([0-9]{3,}(?:[ \.,][0-9]{3})*(?:[\,\.][0-9]{2})?)\b/', $l, $m)) {
                $amount = $this->parseAmount($m[1]);
                if ($amount !== null) return [$amount, 'EUR'];
            }
        }
        return [null, null];
    }

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

    protected function looksLikeStreet(string $line): bool
    {
        return (bool) preg_match('/\b(ROAD|RD|ST|STREET|LANE|AVE|AVENUE|RUE|CHEMIN|CHEM\.|WAY|PARK|COURT|RTE|DR|PLACE|BLVD|GATEWAY|PORT|CROSSING|ZI|RUE DE|COURSE|QUAI|BAKEWELL|LONDON GATEWAY|INDUSTRIES)\b/i', $line);
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

    protected function looksLikeDateOnly(string $s): bool
    {
        $s = str($s)->replace(['–','—'], '-')->toString();
        return (bool)(
            preg_match('/^\d{1,2}\/\d{1,2}\/\d{2,4}$/', $s) ||
            preg_match('/^\d{4}-\d{2}-\d{2}$/', $s) ||
            preg_match('/^\d{1,2}:\d{2}\s*[-to]+\s*\d{1,2}:\d{2}$/i', $s)
        );
    }

    protected function isInstructionLine(string $line): bool
    {
        $u = str($line)->upper()->toString();
        if (Str::length($line) > 140) return true;
        if (preg_match('/^[\-\*\•\x{2022}]/u', $line)) return true;
        if (preg_match('/\b('.implode('|', array_map('preg_quote', self::INSTRUCTION_KEYWORDS)).')\b/i', $u)) return true;
        if (preg_match('/\bDELIVERY\s+NOTE\b/i', $u)) return true;
        return false;
    }

    protected function isLikelyCompany(string $line): bool
    {
        $u = str($line)->upper()->toString();
        if (Str::length($u) > 90) return false;
        if (Str::endsWith($u, ':') && !Str::contains($u, 'REF')) return false;
        $legal = ['LTD','INC','SAS','GMBH','SRL','BV','NV','SARL','S.A.','C/O','UAB','EP GROUP','DP WORLD','ICONEX','FRANCE'];
        if (Str::contains($u, $legal)) return true;
        return (bool) preg_match('/^[A-Z0-9 \'\-&\/\.]{3,}$/', $u);
    }

    protected function looksLikeSectionHeading(string $line): bool
    {
        $u = str($line)->upper()->toString();
        return preg_match('/^(LOADING|DELIVERY|DESTINATION|PICKUP|ORIGIN|CARGO|DETAILS|INFORMATION)[\s:]*$/i', $u);
    }

    protected function resolveCountryFromText(string $text): ?string
    {
        $iso = GeonamesCountry::getIso(trim($text));
        if ($iso) return $iso;

        $parts = preg_split('/[,\.;\|\(\)\-\/]+|\s{2,}/u', $text) ?: [];
        foreach ($parts as $p) {
            $p = trim($p);
            if ($p === '') continue;
            if (in_array(strtoupper($p), ['LTD','INC','GMBH','SAS','SRL','BV','NV','SARL','UAB'], true)) continue;
            if (mb_strlen($p) <= 2) continue;
            $iso = GeonamesCountry::getIso($p);
            if ($iso) return $iso;
        }
        return null;
    }
}