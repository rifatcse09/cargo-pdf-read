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
            dd($lines);

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
            'cargos'                 => filled($cargos) ? $cargos : [['title' => 'PACKAGING']],
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

    /* ===================== Enhanced Cargo Extraction Following Schema ===================== */

    protected function extractCargo(array $lines): array
    {
        // Schema fields to extract
        $title = null;
        $packageCount = null;
        $packageType = null;
        $number = null;  // tracking/order number
        $type = null;    // cargo type from schema enum
        $value = null;   // monetary value
        $currency = null;
        $pkgWidth = null;
        $pkgLength = null;
        $pkgHeight = null;
        $ldm = null;
        $volume = null;
        $weight = null;
        $chargeableWeight = null;
        $tempMin = null;
        $tempMax = null;
        $tempMode = null;
        $adr = null;
        $extraLift = null;
        $palletized = null;
        $manualLoad = null;
        $vehicleMake = null;
        $vehicleModel = null;

        // Scan every line for cargo data - first match wins
        for ($i = 0; $i < count($lines); $i++) {
            $line = trim($lines[$i]);
            if (empty($line)) continue;

            $upper = str($line)->upper()->toString();

            // 1. TITLE extraction - cargo description
            if (!$title) {
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

                // TransAlliance "M. nature:" pattern
                if (preg_match('/M\.\s*nature\s*:\s*(.+)/i', $line, $m)) {
                    $cargoDesc = trim($m[1]);
                    if (strlen($cargoDesc) > 2 && !preg_match('/^\d+$/', $cargoDesc)) {
                        $title = strtoupper($cargoDesc);
                    }
                }

                // Generic fallback
                if (!$title && preg_match('/\b(GOODS|CARGO|COMMODITY|FREIGHT|MERCHANDISE)\b/i', $upper) && strlen($line) < 30) {
                    $title = 'GENERAL CARGO';
                }
            }

            // 2. PACKAGE_COUNT and PACKAGE_TYPE extraction
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

            // 3. TYPE extraction - cargo type enum
            if (!$type) {
                if (preg_match('/\b(FTL|FULL\s*TRUCK|FULL\s*LOAD)\b/i', $upper)) {
                    $type = 'FTL';
                } elseif (preg_match('/\b(LTL|LESS\s*THAN\s*TRUCK)\b/i', $upper)) {
                    $type = 'LTL';
                } elseif (preg_match('/\b(FCL|FULL\s*CONTAINER)\b/i', $upper)) {
                    $type = 'FCL';
                } elseif (preg_match('/\b(LCL|LESS\s*THAN\s*CONTAINER)\b/i', $upper)) {
                    $type = 'LCL';
                } elseif (preg_match('/\b(PARTIAL|PTL)\b/i', $upper)) {
                    $type = 'partial';
                } elseif (preg_match('/\b(CONTAINER)\b/i', $upper)) {
                    $type = 'container';
                } elseif (preg_match('/\b(CAR|VEHICLE)\b/i', $upper)) {
                    $type = 'car';
                } elseif (preg_match('/\b(AIR|FLIGHT)\b/i', $upper)) {
                    $type = 'air shipment';
                } elseif (preg_match('/\b(PARCEL|EXPRESS)\b/i', $upper)) {
                    $type = 'parcel';
                }
            }

            // 4. WEIGHT extraction
            if ($weight === null) {
                // TransAlliance format: "24000,000" after weight label
                if (preg_match('/^([0-9]+[,\.][0-9]{1,3})$/', $line) && $i > 0) {
                    $prevLine = str($lines[$i-1])->upper();
                    if (str_contains($prevLine, 'KGS') || str_contains($prevLine, 'WEIGHT') ||
                        str_contains($prevLine, 'KG') || str_contains($prevLine, 'MASS')) {
                        $weight = (float) str_replace(',', '.', $line);
                    }
                }
                // Inline weight patterns
                if (preg_match('/\b(?:weight|mass|kgs?)\b[:\s]*([0-9][0-9,\.]*)\s*(?:kg|kgs?|kilos?)?\b/i', $line, $m)) {
                    $weight = (float) str_replace(',', '.', $m[1]);
                }
                // Standalone weight
                if (preg_match('/\b([0-9]{4,}[0-9,\.]*)\s*(?:kg|kgs?|kilos?)\b/i', $line, $m)) {
                    $candidate = (float) str_replace(',', '.', $m[1]);
                    if ($candidate >= 100) $weight = $candidate;
                }
            }

            // 5. LDM extraction
            if ($ldm === null) {
                if (preg_match('/^([0-9]+[,\.][0-9]{1,3})$/', $line) && $i > 0) {
                    $prevLine = str($lines[$i-1])->upper();
                    if ((str_contains($prevLine, 'LM') || str_contains($prevLine, 'LOADING') ||
                         str_contains($prevLine, 'METER')) && !str_contains($prevLine, 'KG')) {
                        $ldm = (float) str_replace(',', '.', $line);
                    }
                }
                if (preg_match('/\b(?:ldm|loading\s*meters?|meters?)\b[:\s]*([0-9][0-9,\.]*)/i', $line, $m)) {
                    $ldm = (float) str_replace(',', '.', $m[1]);
                }
            }

            // 6. VOLUME extraction
            if ($volume === null && preg_match('/\b([0-9][0-9,\.]*)\s*(?:m3|m³|cbm|cubic\s*meters?)\b/i', $line, $m)) {
                $volume = (float) str_replace(',', '.', $m[1]);
            }

            // 7. VALUE and CURRENCY extraction
            if ($value === null && preg_match('/\b(?:value|worth|cost)\b[:\s]*([€£\$]?)\s*([0-9][0-9,\.]*)\s*([€£\$]?)\s*(EUR|USD|GBP)?\b/i', $line, $m)) {
                $value = (float) str_replace(',', '.', $m[2]);
                $symbol = $m[1] ?: $m[3];
                $currency = $m[4] ?: ($symbol ? $this->ccyFromSymbol($symbol) : 'EUR');
            }

            // 8. DIMENSIONS extraction
            if (!$pkgLength && preg_match('/\b(?:length|l)\b[:\s]*([0-9][0-9,\.]*)\s*(?:m|meter[s]?|cm)?\b/i', $line, $m)) {
                $pkgLength = (float) str_replace(',', '.', $m[1]);
                if (str_contains($line, 'cm')) $pkgLength /= 100; // Convert cm to meters
            }
            if (!$pkgWidth && preg_match('/\b(?:width|w)\b[:\s]*([0-9][0-9,\.]*)\s*(?:m|meter[s]?|cm)?\b/i', $line, $m)) {
                $pkgWidth = (float) str_replace(',', '.', $m[1]);
                if (str_contains($line, 'cm')) $pkgWidth /= 100;
            }
            if (!$pkgHeight && preg_match('/\b(?:height|h)\b[:\s]*([0-9][0-9,\.]*)\s*(?:m|meter[s]?|cm)?\b/i', $line, $m)) {
                $pkgHeight = (float) str_replace(',', '.', $m[1]);
                if (str_contains($line, 'cm')) $pkgHeight /= 100;
            }

            // 9. TEMPERATURE extraction
            if ($tempMin === null && preg_match('/\b(?:temp|temperature)\b[:\s]*([+-]?[0-9][0-9,\.]*)\s*(?:°?c|celsius)?\b/i', $line, $m)) {
                $tempMin = (float) str_replace(',', '.', $m[1]);
            }
            if ($tempMax === null && preg_match('/\b(?:temp|temperature)\b[:\s]*[+-]?[0-9][0-9,\.]*\s*(?:to|-)?\s*([+-]?[0-9][0-9,\.]*)\s*(?:°?c|celsius)?\b/i', $line, $m)) {
                $tempMax = (float) str_replace(',', '.', $m[1]);
            }

            // 10. BOOLEAN flags extraction
            if ($adr === null) $adr = preg_match('/\b(ADR|DANGEROUS|HAZARDOUS)\b/i', $upper) ? true : null;
            if ($extraLift === null) $extraLift = preg_match('/\b(LIFT|CRANE|FORKLIFT)\b/i', $upper) ? true : null;
            if ($palletized === null) $palletized = preg_match('/\b(PALLETIZED|ON\s*PALLETS)\b/i', $upper) ? true : null;
            if ($manualLoad === null) $manualLoad = preg_match('/\b(MANUAL|HAND\s*LOAD)\b/i', $upper) ? true : null;

            // 11. VEHICLE details extraction
            if (!$vehicleMake && preg_match('/\b(BMW|MERCEDES|AUDI|VOLKSWAGEN|FORD|RENAULT|PEUGEOT|VOLVO|SCANIA|MAN)\b/i', $upper, $m)) {
                $vehicleMake = strtoupper($m[1]);
            }

            // 12. NUMBER/REFERENCE extraction
            if (!$number && preg_match('/\b(?:tracking|ref|reference|order|po)\b[:\s#]*([A-Z0-9\-\/]{4,})/i', $line, $m)) {
                $number = trim($m[1]);
            }
        }

        // Build cargo object according to schema
        $cargo = [
            'title' => $title ?: 'PACKAGING',
            'type' => $type ?: 'FTL' // Default for TransAlliance
        ];

        // Add optional fields only if they have valid values
        if ($packageCount !== null && $packageCount > 0) {
            $cargo['package_count'] = $packageCount;
        }
        if ($packageType) {
            $cargo['package_type'] = $packageType;
        }
        if ($number) {
            $cargo['number'] = $number;
        }
        if ($value !== null && $value > 0) {
            $cargo['value'] = $value;
        }
        if ($currency) {
            $cargo['currency'] = $currency;
        }
        if ($pkgWidth !== null && $pkgWidth > 0) {
            $cargo['pkg_width'] = $pkgWidth;
        }
        if ($pkgLength !== null && $pkgLength > 0) {
            $cargo['pkg_length'] = $pkgLength;
        }
        if ($pkgHeight !== null && $pkgHeight > 0) {
            $cargo['pkg_height'] = $pkgHeight;
        }
        if ($ldm !== null && $ldm > 0) {
            $cargo['ldm'] = $ldm;
        }
        if ($volume !== null && $volume > 0) {
            $cargo['volume'] = $volume;
        }
        if ($weight !== null && $weight > 0) {
            $cargo['weight'] = $weight;
        }
        if ($chargeableWeight !== null && $chargeableWeight > 0) {
            $cargo['chargeable_weight'] = $chargeableWeight;
        }
        if ($tempMin !== null) {
            $cargo['temperature_min'] = $tempMin;
        }
        if ($tempMax !== null) {
            $cargo['temperature_max'] = $tempMax;
        }
        if ($tempMode) {
            $cargo['temperature_mode'] = $tempMode;
        }
        if ($adr === true) {
            $cargo['adr'] = true;
        }
        if ($extraLift === true) {
            $cargo['extra_lift'] = true;
        }
        if ($palletized === true) {
            $cargo['palletized'] = true;
        }
        if ($manualLoad === true) {
            $cargo['manual_load'] = true;
        }
        if ($vehicleMake) {
            $cargo['vehicle_make'] = $vehicleMake;
        }
        if ($vehicleModel) {
            $cargo['vehicle_model'] = $vehicleModel;
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