<?php

namespace App\Assistants;

use Carbon\Carbon;
use App\GeonamesCountry;
use Illuminate\Support\Arr;
use Illuminate\Support\Str;

class ZieglerPdfAssistant extends PdfClient
{
    public static function validateFormat(array $lines)
    {
        $head = collect($lines)->take(50)->map(fn($l) => Str::upper(trim($l)))->implode(' ');

        $hasZiegler = Str::contains($head, 'ZIEGLER UK LTD');
        $hasBooking = Str::contains($head, ['BOOKING INSTRUCTION', 'BOOKING', 'ORDER']);

        return $hasZiegler && $hasBooking;
    }

    public function processLines(array $lines, ?string $attachment_filename = null)
    {
        // Clean and normalize lines
        $lines = collect($lines)
            ->map(fn($l) => trim(preg_replace('/\s+/', ' ', $l)))
            ->filter(fn($l) => !empty($l))
            ->values()
            ->toArray();
            //dd( $lines);

        // Extract all components
        $customer = $this->extractCustomer($lines);
        $orderRef = $this->extractOrderReference($lines);
        $freight = $this->extractFreight($lines);
        $loadingLocs = $this->extractLoadingLocations($lines);
        $destLocs = $this->extractDestinationLocations($lines);
        $cargos = $this->extractCargos($lines);
        $comment = $this->extractComment($lines);

        // Ensure we have proper fallback locations if none found
        if (empty($loadingLocs)) {
            $loadingLocs = [[
                'company_address' => [
                    'company' => 'Collection Location',
                    'street_address' => 'TBC',
                    'city' => 'Unknown',
                    'postal_code' => 'TBC',
                    'country_code' => 'GB'
                ]
            ]];
        }

        if (empty($destLocs)) {
            $destLocs = [[
                'company_address' => [
                    'company' => 'Delivery Location',
                    'street_address' => 'TBC',
                    'city' => 'Unknown',
                    'postal_code' => 'TBC',
                    'country_code' => 'GB'
                ]
            ]];
        }

        // Build payload according to schema
        $payload = [
            'attachment_filenames' => $attachment_filename ? [basename($attachment_filename)] : [],
            'customer' => $customer,
            'order_reference' => $orderRef,
            'freight_price' => $freight['price'],
            'freight_currency' => $freight['currency'] ?: 'EUR',
            'loading_locations' => $loadingLocs,
            'destination_locations' => $destLocs,
            'cargos' => $cargos ?: [['title' => 'General cargo', 'package_count' => 1]],
            'comment' => $comment,
        ];

        $this->createOrder($payload);
        return $payload;
    }

    protected function extractCustomer(array $lines): array
    {
        $details = [];

        // Find Ziegler header block
        $headerIdx = collect($lines)->search(fn($line) =>
            Str::contains(Str::upper($line), 'ZIEGLER UK LTD')
        );

        if ($headerIdx !== false) {
            $details['company'] = 'ZIEGLER UK LTD';

            // Extract address from the actual PDF structure
            $addressLines = array_slice($lines, $headerIdx + 1, 6);

            // Use the parseZieglerAddress method for proper country detection
            $this->parseZieglerAddress($addressLines, $details);

            // Fallback values based on known Ziegler address if not extracted
            if (!isset($details['street_address'])) {
                $details['street_address'] = 'LONDON GATEWAY LOGISTICS PARK, NORTH 4, NORTH SEA CROSSING';
            }
            if (!isset($details['city'])) {
                $details['city'] = 'STANFORD LE HOPE';
            }
            if (!isset($details['postal_code'])) {
                $details['postal_code'] = 'SS17 9FJ';
                if (!isset($details['country_code'])) {
                    $details['country_code'] = 'GB';
                }
            }
            if (!isset($details['country_code'])) {
                $details['country_code'] = 'GB';
            }
        }

        // Find customer comment from terms section
        $comment = $this->findZieglerTerms($lines);
        if ($comment) {
            $details['comment'] = $comment;
        }

        return [
            'side' => 'none',
            'details' => $details
        ];
    }

    protected function parseZieglerAddress(array $addressLines, array &$details): void
    {
        $streetParts = [];
        $cityFound = false;
        $postalCodeFound = false;

        foreach ($addressLines as $lineIndex => $line) {
            $line = trim($line);
            if (empty($line)) break;

            // Stop at booking/instruction lines
            if (preg_match('/\b(BOOKING|INSTRUCTION|TELEPHONE|REF)\b/i', $line)) {
                break;
            }

            // Check for UK postcode first (SS17 9FJ)
            if (preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2})\b/', $line, $m)) {
                $pc = strtoupper(str_replace(' ', '', $m[1]));
                $details['postal_code'] = substr($pc, 0, -3) . ' ' . substr($pc, -3);
                $postalCodeFound = true;

                if (!empty($details['postal_code'])) {
                    $countryIso = GeonamesCountry::getIso($details['postal_code']);
                    if ($countryIso) {
                        $details['country_code'] = $countryIso;
                    } else {
                        $details['country_code'] = 'GB';
                    }
                } else {
                    $details['country_code'] = 'GB';
                }
                continue;
            }

            // Check for French/EU postcode pattern (5 digits)
            if (preg_match('/\b(\d{5})\b/', $line, $m)) {
                $details['postal_code'] = $m[1];
                $postalCodeFound = true;

                if (!empty($details['postal_code'])) {
                    $countryIso = GeonamesCountry::getIso($details['postal_code']);
                    if ($countryIso) {
                        $details['country_code'] = $countryIso;
                    } else {
                        $details['country_code'] = 'FR';
                    }
                } else {
                    $details['country_code'] = 'FR';
                }
                continue;
            }

            // Check if this is a city line
            if (!$cityFound && preg_match('/^[A-Z\s\'-]{4,}$/', $line) && !preg_match('/\d/', $line)) {
                $isLikelyCity = false;

                if ($postalCodeFound) {
                    $isLikelyCity = true;
                }

                // Check if next line has postal code
                if (!$isLikelyCity && isset($addressLines[$lineIndex + 1])) {
                    $nextLine = trim($addressLines[$lineIndex + 1]);
                    if (preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2}|\d{5})\b/', $nextLine)) {
                        $isLikelyCity = true;
                    }
                }

                // Specific city patterns for known cities
                if (!$isLikelyCity && preg_match('/^(STANFORD LE HOPE|LONDON|BIRMINGHAM|MANCHESTER|BRISTOL|CARDIFF|GLASGOW|EDINBURGH|BELFAST)$/i', $line)) {
                    $isLikelyCity = true;
                }

                if ($isLikelyCity) {
                    $details['city'] = $line;
                    $cityFound = true;
                    continue;
                }
            }

            $streetParts[] = $line;
        }

        // Build the complete street address
        if (!empty($streetParts)) {
            $details['street_address'] = collect($streetParts)->implode(', ');
        }
    }

    protected function findZieglerTerms(array $lines): ?string
    {
        $parts = [];

        foreach ($lines as $line) {
            $lowerLine = Str::lower($line);

            if (Str::contains($lowerLine, ['purchase order general terms', 'purchase order terms'])) {
                $parts[] = 'Ziegler purchase order terms apply';
            }

            if (Str::contains($lowerLine, ['please quote our reference number', 'quote our reference on your invoice'])) {
                $parts[] = 'Please quote our reference on invoice';
            }
        }

        if (empty($parts)) {
            return null;
        }

        return collect($parts)->unique()->implode('. ') . '.';
    }

    protected function extractOrderReference(array $lines): ?string
    {
        $foundReferences = [];

        foreach ($lines as $i => $line) {
            // "Ziegler Ref XXXXX" - highest priority - capture any reference dynamically
            if (preg_match('/Ziegler\s+Ref\s+([A-Z0-9\/\-]+)/i', $line, $m)) {
                $foundReferences[] = ['priority' => 1, 'ref' => $m[1], 'line' => $i, 'pattern' => 'Ziegler Ref'];
            }

            // "Ziegler Ref" in current line, reference number in next line
            if (preg_match('/^Ziegler\s+Ref$/i', trim($line)) && isset($lines[$i + 1])) {
                $nextLine = trim($lines[$i + 1]);
                if (preg_match('/^([A-Z0-9\/\-]+)$/i', $nextLine, $m)) {
                    $foundReferences[] = ['priority' => 1, 'ref' => $m[1], 'line' => $i, 'pattern' => 'Ziegler Ref (next line)'];
                }
            }

            // "Our Ref: XXXXXXX" - high priority
            if (preg_match('/Our\s+Ref(?:erence)?[:\s]+([A-Z0-9\-\/]+)/i', $line, $m)) {
                $foundReferences[] = ['priority' => 2, 'ref' => trim($m[1]), 'line' => $i, 'pattern' => 'Our Ref'];
            }

            // "Ref: XXXXXXX" at start of line - medium priority
            if (preg_match('/^Ref(?:erence)?[:\s]+([A-Z0-9\-\/]+)/i', $line, $m)) {
                $foundReferences[] = ['priority' => 3, 'ref' => trim($m[1]), 'line' => $i, 'pattern' => 'Line Ref'];
            }

            // Generic "Ref" followed by alphanumeric - lower priority
            if (preg_match('/\bRef[:\s]+([A-Z0-9\-\/]{4,})/i', $line, $m)) {
                $ref = trim($m[1]);
                if (!preg_match('/^(ZIEGLER|LONDON|GATEWAY|BOOKING)$/i', $ref)) {
                    $foundReferences[] = ['priority' => 4, 'ref' => $ref, 'line' => $i, 'pattern' => 'Generic Ref'];
                }
            }

            // Order/Booking patterns - lower priority
            if (preg_match('/(?:Order|Booking)(?:\s+(?:No|Number|Ref|Reference))?[:#\s]*([A-Z0-9\-\/]{4,})/i', $line, $m)) {
                $ref = trim($m[1]);
                if (!preg_match('/^(INSTRUCTION|ZIEGLER)$/i', $ref)) {
                    $foundReferences[] = ['priority' => 5, 'ref' => $ref, 'line' => $i, 'pattern' => 'Order/Booking'];
                }
            }
        }

        // Sort by priority (lower number = higher priority)
        usort($foundReferences, function($a, $b) {
            return $a['priority'] <=> $b['priority'];
        });

        // Return the highest priority reference
        if (!empty($foundReferences)) {
            return $foundReferences[0]['ref'];
        }

        return null;
    }

    protected function extractFreight(array $lines): array
    {
        $price = null;
        $currency = null;

        foreach ($lines as $line) {
            // "Rate € 1,000" or "Rate € 1.15" - more flexible pattern
            if (preg_match('/Rate\s*€\s*([0-9,.\s]+)/i', $line, $m)) {
                $cleanPrice = function_exists('uncomma') ? uncomma($m[1]) : str_replace([',', ' '], '', $m[1]);
                $price = (float) $cleanPrice;
                $currency = 'EUR';
                break;
            }

            // "€ 1,000" or "€ 1.15" - Euro symbol first
            if (preg_match('/€\s*([0-9,.\s]+)/i', $line, $m)) {
                $cleanPrice = function_exists('uncomma') ? uncomma($m[1]) : str_replace([',', ' '], '', $m[1]);
                $price = (float) $cleanPrice;
                $currency = 'EUR';
                break;
            }

            // "1,000 €" or "1.15 €" - Euro symbol after
            if (preg_match('/([0-9,.\s]+)\s*€/i', $line, $m)) {
                $cleanPrice = function_exists('uncomma') ? uncomma($m[1]) : str_replace([',', ' '], '', $m[1]);
                $price = (float) $cleanPrice;
                $currency = 'EUR';
                break;
            }

            // "Price: EUR 1000"
            if (preg_match('/Price[:\s]*EUR\s*([0-9,.\s]+)/i', $line, $m)) {
                $cleanPrice = function_exists('uncomma') ? uncomma($m[1]) : str_replace([',', ' '], '', $m[1]);
                $price = (float) $cleanPrice;
                $currency = 'EUR';
                break;
            }

            // Generic currency patterns
            if (preg_match('/(EUR|GBP|USD)\s*([0-9,.\s]+)/i', $line, $m)) {
                $currency = strtoupper($m[1]);
                $cleanPrice = function_exists('uncomma') ? uncomma($m[2]) : str_replace([',', ' '], '', $m[2]);
                $price = (float) $cleanPrice;
                break;
            }

            if (preg_match('/([€£$])\s*([0-9,.\s]+)/', $line, $m)) {
                $currency = ['€' => 'EUR', '£' => 'GBP', '$' => 'USD'][$m[1]] ?? null;
                $cleanPrice = function_exists('uncomma') ? uncomma($m[2]) : str_replace([',', ' '], '', $m[2]);
                $price = (float) $cleanPrice;
                if ($currency) break;
            }

            // Look for just currency symbols/codes without price
            if (preg_match('/\b(EUR|GBP|USD)\b/i', $line, $m)) {
                $currency = $currency ?: strtoupper($m[1]);
            }
            if (Str::contains($line, '€') && !$currency) {
                $currency = 'EUR';
            }
            if (Str::contains($line, '£') && !$currency) {
                $currency = 'GBP';
            }
        }

        return [
            'price' => $price,
            'currency' => $currency ?: 'EUR'
        ];
    }

    protected function extractLoadingLocations(array $lines): array
    {
        $locations = [];
        $processedSections = [];

        $collectionStarts = $this->findCollectionSectionStarts($lines);

        if (empty($collectionStarts)) {
            $collectionStarts = $this->findCollectionStartsAggressive($lines);
        }

        usort($collectionStarts, function($a, $b) {
            if ($a['confidence'] !== $b['confidence']) {
                return $b['confidence'] <=> $a['confidence'];
            }
            return $a['line'] <=> $b['line'];
        });

        $maxLocations = 2;
        foreach ($collectionStarts as $start) {
            if (count($locations) >= $maxLocations) {
                break;
            }

            $startLine = $start['line'];

            $shouldSkip = false;
            foreach ($processedSections as $processed) {
                if (abs($startLine - $processed['start']) < 8) {
                    $shouldSkip = true;
                    break;
                }
            }

            if ($shouldSkip) continue;

            $endLine = $this->findCollectionSectionEnd($lines, $startLine);
            $location = $this->parseCollectionLocationSection($lines, $startLine, $endLine);

            if ($location) {
                $isDuplicate = false;
                foreach ($locations as $existingLocation) {
                    if ($this->areLocationsDuplicate($location, $existingLocation)) {
                        $isDuplicate = true;
                        break;
                    }
                }

                if (!$isDuplicate) {
                    $locations[] = $location;
                    $processedSections[] = [
                        'start' => $startLine,
                        'end' => $endLine,
                        'company' => $location['company_address']['company']
                    ];
                }
            }
        }

        return $locations;
    }

    protected function extractDestinationLocations(array $lines): array
    {
        $locations = [];

        // Strategy 1: Look for explicit "Delivery" patterns
        foreach ($lines as $i => $line) {
            if (preg_match('/^Delivery\s+(.+)/i', $line, $m)) {
                $location = $this->parseDeliveryLocationBlock($lines, $i);
                if ($location) {
                    $locations[] = $location;
                }
            }
        }

        // Strategy 2: Look for company names that appear to be destinations
        $destinationCompanyPatterns = [
            '/^(ICD8)$/i',
            '/^(AKZO NOBEL C\/O [A-Z]+)$/i',
            '/^(DSL\s*\([^)]+\))$/i',
            '/^([A-Z][A-Z\s&\(\)\/\-\.C\/O]+)\s+REF$/i'
        ];

        for ($i = 0; $i < count($lines); $i++) {
            $line = trim($lines[$i]);
            $foundPattern = false;
            $matchingLine = $i;

            foreach ($destinationCompanyPatterns as $pattern) {
                if (preg_match($pattern, $line, $m)) {
                    $foundPattern = true;
                    $matchingLine = $i;
                    break;
                }
            }

            if (!$foundPattern && isset($lines[$i + 1])) {
                $nextLine = trim($lines[$i + 1]);
                if (preg_match('/^(DSL\s*\([^)]+\))$/i', $nextLine, $m)) {
                    $foundPattern = true;
                    $matchingLine = $i;
                }
            }

            if ($foundPattern) {
                $isDeliverySection = $this->isLikelyDeliverySection($lines, $matchingLine);

                if ($isDeliverySection) {
                    $location = $this->parseDeliveryLocationBlock($lines, $matchingLine);
                    if ($location) {
                        $isDupe = false;
                        foreach ($locations as $existing) {
                            if ($existing['company_address']['company'] === $location['company_address']['company']) {
                                $isDupe = true;
                                break;
                            }
                        }
                        if (!$isDupe) {
                            $locations[] = $location;
                        }
                    }
                }
            }
        }

        return $locations;
    }

    protected function isLikelyDeliverySection(array $lines, int $startIdx): bool
    {
        $endIdx = min($startIdx + 8, count($lines) - 1);
        $sectionLines = array_slice($lines, $startIdx, $endIdx - $startIdx + 1);

        $hasAddressInfo = false;
        $hasTimeInfo = false;
        $hasRefInfo = false;

        foreach ($sectionLines as $line) {
            if (preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2}|\d{5})\b/', $line) ||
                preg_match('/\b(RUE|ROAD|STREET|AVENUE|CHEMIN|Chem\.)\b/i', $line)) {
                $hasAddressInfo = true;
            }

            if (preg_match('/\d{2}:\d{2}|BOOKED|Time/i', $line)) {
                $hasTimeInfo = true;
            }

            if (preg_match('/REF\s+[A-Z0-9]+/i', $line)) {
                $hasRefInfo = true;
            }
        }

        return $hasAddressInfo && ($hasTimeInfo || $hasRefInfo);
    }

    protected function parseDeliveryLocationBlock(array $lines, int $startIdx): ?array
    {
        $blockLines = [];
        $maxLookAhead = 15;

        for ($i = $startIdx; $i < count($lines) && $i < $startIdx + $maxLookAhead; $i++) {
            $line = trim($lines[$i]);
            if (empty($line)) continue;

            if ($i > $startIdx && preg_match('/^(Collection|Delivery|Rate\s*€|Terms|Payment|Carrier)/i', $line)) {
                break;
            }

            if ($i > $startIdx + 3 && preg_match('/^[A-Z][A-Z\s&\(\)\/\-\.]+\s+REF\s+[A-Z0-9\/\-]+$/i', $line) &&
                !preg_match('/\b(ROAD|STREET|RUE|AVENUE|CHEMIN)\b/i', $line)) {
                break;
            }

            $blockLines[] = $line;
        }

        if (empty($blockLines)) {
            return null;
        }

        $company = '';

        $searchStartIdx = max(0, $startIdx - 5);
        $searchEndIdx = min(count($lines) - 1, $startIdx + 25);

        for ($i = $searchStartIdx; $i <= $searchEndIdx; $i++) {
            $line = trim($lines[$i]);

            if (preg_match('/^DSL\s*\([^)]+\)$/i', $line, $m)) {
                $company = trim($line);
                break;
            }

            if (preg_match('/\bDSL\s*\([^)]+\)/i', $line, $m)) {
                $company = trim($m[0]);
                break;
            }
        }

        if (empty($company)) {
            for ($i = $searchStartIdx; $i <= $searchEndIdx; $i++) {
                $line = trim($lines[$i]);

                if (preg_match('/\b(delivery\s+slot|will\s+be\s+provided|soon|REF|ROAD|STREET|RUE|AVENUE|CHEMIN|\d{2}\/\d{2}\/\d{4})\b/i', $line)) {
                    continue;
                }

                if (preg_match('/^Delivery\s+(.+)$/i', $line, $m)) {
                    $potentialCompany = trim($m[1]);
                    if (!preg_match('/\b(slot|will|be|provided|soon)\b/i', $potentialCompany)) {
                        $company = $potentialCompany;
                        break;
                    }
                }

                if (preg_match('/^(ICD8|AKZO NOBEL C\/O [A-Z]+)$/i', $line, $m)) {
                    $company = trim($m[1]);
                    break;
                }
            }
        }

        $address = [
            'company' => $company ?: 'Delivery Location',
            'street_address' => '',
            'city' => '',
            'postal_code' => '',
            'country_code' => 'FR'
        ];

        $timeObj = null;
        $dateFound = null;
        $timeStart = null;
        $timeEnd = null;

        for ($i = $searchStartIdx; $i <= $searchEndIdx; $i++) {
            $line = trim($lines[$i]);

            if (!empty($company) && (
                trim($line) === trim($company) ||
                preg_match('/^Delivery\s+/i', $line)
            )) {
                continue;
            }

            if (preg_match('/(\d{2}\/\d{2}\/\d{4})/', $line, $m)) {
                $dateFound = $m[1];
            }

            if (preg_match('/BOOKED[-\s]*(\d{1,2}):?(\d{2})?\s*(AM|PM)?/i', $line, $m)) {
                $hour = (int)$m[1];
                $minute = isset($m[2]) ? (int)$m[2] : 0;
                $ampm = isset($m[3]) ? strtolower($m[3]) : '';

                if ($ampm === 'pm' && $hour < 12) {
                    $hour += 12;
                } elseif ($ampm === 'am' && $hour === 12) {
                    $hour = 0;
                }

                $timeStart = sprintf('%02d:%02d', $hour, $minute);
            }
            elseif (preg_match('/(\d{1,2}):(\d{2})\s+Time\s+To:\s+(\d{1,2}):(\d{2})/i', $line, $m)) {
                $timeStart = sprintf('%02d:%02d', (int)$m[1], (int)$m[2]);
                $timeEnd = sprintf('%02d:%02d', (int)$m[3], (int)$m[4]);
            }
            elseif (preg_match('/(\d{2}):(\d{2})\s*[-–]\s*(\d{2}):(\d{2})/', $line, $m)) {
                $timeStart = sprintf('%02d:%02d', (int)$m[1], (int)$m[2]);
                $timeEnd = sprintf('%02d:%02d', (int)$m[3], (int)$m[4]);
            }

            $this->parseDeliveryAddressComponents($line, $address);
        }

        if ($dateFound) {
            try {
                $carbonDate = Carbon::createFromFormat('d/m/Y', $dateFound);

                if ($timeStart) {
                    $timeObj = [
                        'datetime_from' => $carbonDate->copy()->setTimeFromTimeString($timeStart)->format('c')
                    ];

                    if ($timeEnd) {
                        $timeObj['datetime_to'] = $carbonDate->copy()->setTimeFromTimeString($timeEnd)->format('c');
                    }
                } else {
                    $timeObj = [
                        'datetime_from' => $carbonDate->copy()->setTime(8, 0)->format('c')
                    ];
                }
            } catch (\Exception $e) {
                // Continue without time object
            }
        }

        if (empty($address['street_address'])) {
            $address['street_address'] = 'TBC';
        }
        if (empty($address['city'])) {
            $address['city'] = 'Unknown';
        }
        if (empty($address['postal_code'])) {
            $address['postal_code'] = 'TBC';
        }

        $result = ['company_address' => $address];
        if ($timeObj) {
            $result['time'] = $timeObj;
        }

        return $result;
    }

    protected function parseDeliveryAddressComponents(string $line, array &$address): void
    {
        if (preg_match('/^REF\s+[A-Z0-9]+/i', $line) ||
            preg_match('/^(ICD8|AKZO NOBEL|DSL|WH\s)/i', $line) ||
            preg_match('/BOOKED/i', $line) ||
            preg_match('/Time\s+To:/i', $line) ||
            preg_match('/\d{2}\/\d{2}\/\d{4}/', $line) ||
            preg_match('/\b(delivery\s+slot\s+will\s+be\s+provided|slot\s+will\s+be\s+provided|soon)\b/i', $line) ||
            preg_match('/^\d+\s+(pallets?|packages?)/i', $line)) {
            return;
        }

        if (preg_match('/^Delivery$/i', $line)) {
            return;
        }

        // Enhanced French address pattern: "STIRING WENDEL, FR57350"
        if (preg_match('/^([A-Z\s\'-]+),\s*FR(\d{5})$/i', $line, $m)) {
            $address['city'] = trim($m[1]);
            $address['postal_code'] = $m[2];
            $address['country_code'] = 'FR';
            return;
        }

        // Standard French postcode + city pattern: "57350 Stiring Wendel" or "45770 SARAN"
        if (preg_match('/^(\d{5})\s+(.+)$/i', $line, $m)) {
            $address['postal_code'] = $m[1];
            $address['city'] = trim($m[2]);
            $address['country_code'] = 'FR';
            return;
        }

        // Enhanced street address patterns for French addresses
        if (preg_match('/^(\d+\s+(?:RUE|AVENUE|BOULEVARD|CHEMIN|Chem\.)\s+.+)$/i', $line, $m)) {
            $address['street_address'] = trim($m[1]);
            $address['country_code'] = 'FR';
            return;
        }

        if (preg_match('/^((?:RUE|AVENUE|BOULEVARD|CHEMIN|Chem\.)\s+.+)$/i', $line, $m)) {
            $address['street_address'] = trim($m[1]);
            $address['country_code'] = 'FR';
            return;
        }

        if (preg_match('/^(\d+\s+Chem\.\s+.+)$/i', $line, $m)) {
            $address['street_address'] = trim($m[1]);
            $address['country_code'] = 'FR';
            return;
        }

        // UK patterns as fallback
        if (preg_match('/([A-Z]{1,2}\d+\s+\d[A-Z]{2})\s+([A-Z\s]+)$/i', $line, $m)) {
            $address['postal_code'] = $m[1];
            $address['city'] = trim($m[2]);
            $address['country_code'] = 'GB';
            return;
        }

        if (preg_match('/^(.+(?:ROAD|STREET|LANE|AVENUE|DRIVE|CLOSE|WAY).*?)$/i', $line, $m) && empty($address['street_address'])) {
            $address['street_address'] = trim($m[1]);
            return;
        }

        // Standalone postal code
        if (preg_match('/\b(\d{5})\b/', $line, $m) && empty($address['postal_code'])) {
            $address['postal_code'] = $m[1];
            $address['country_code'] = 'FR';
        }
        elseif (preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2})\b/', $line, $m) && empty($address['postal_code'])) {
            $pc = strtoupper(str_replace(' ', '', $m[1]));
            $address['postal_code'] = substr($pc, 0, -3) . ' ' . substr($pc, -3);
            $address['country_code'] = 'GB';
        }

        // Standalone city
        if (preg_match('/^([A-Z\s\'-]{3,30})$/i', $line) &&
            !preg_match('/\d/', $line) &&
            empty($address['city']) &&
            !preg_match('/\b(ROAD|STREET|LANE|AVENUE|WAY|CLOSE|UNITS|HALL|LTD|LIMITED|SOLUTIONS|LOGISTICS|FULFILMENT|REF|COLLECTION|DELIVERY|RUE|CHEMIN|BOULEVARD)\b/i', $line)) {
            $address['city'] = trim($line);
        }
    }

    protected function extractCargos(array $lines): array
    {
        $cargos = [];
        $foundCargoCounts = [];

        $currentSection = null;
        $currentCompany = null;

        for ($i = 0; $i < count($lines); $i++) {
            $line = trim($lines[$i]);

            if (preg_match('/^Collection\s+(.+)/i', $line, $m)) {
                $currentSection = 'collection';
                $currentCompany = trim($m[1]);
                continue;
            }

            if (preg_match('/^Collection$/i', $line)) {
                $currentSection = 'collection';
                $currentCompany = null;
                continue;
            }

            if (preg_match('/^Delivery/i', $line)) {
                $currentSection = 'delivery';
                $currentCompany = null;
                continue;
            }

            if ($currentSection && !$currentCompany) {
                if (preg_match('/^(AKZO NOBEL|EPAC FULFILMENT SOLUTIONS LTD|LINDAL VALVE CO LTD|IBF|ICD8)$/i', $line)) {
                    $currentCompany = trim($line);
                    continue;
                }
            }

            if ($currentSection === 'collection' && $currentCompany) {
                if (preg_match('/(\d+)\s+(?:PALLETS?|PALLET)/i', $line, $m)) {
                    $palletCount = (int) $m[1];
                    $cargoKey = $palletCount . '_pallets';

                    if (!isset($foundCargoCounts[$cargoKey])) {
                        $loadNumber = count($cargos) + 1;
                        $cargos[] = [
                            'title' => "Mixed Pallet Load {$loadNumber}",
                            'package_count' => $palletCount,
                            'package_type' => 'pallet',
                            'type' => 'FTL'
                        ];
                        $foundCargoCounts[$cargoKey] = true;
                    }
                }
                elseif (preg_match('/(\d+)\s+(PACKAGES?|CARTONS?|BOXES?|ITEMS?)/i', $line, $m)) {
                    $packageCount = (int) $m[1];
                    $packageType = strtolower(rtrim($m[2], 's'));
                    $cargoKey = $packageCount . '_' . $packageType;

                    if (!isset($foundCargoCounts[$cargoKey])) {
                        $loadNumber = count($cargos) + 1;
                        $cargos[] = [
                            'title' => "Mixed Package Load {$loadNumber}",
                            'package_count' => $packageCount,
                            'package_type' => $packageType,
                            'type' => 'FTL'
                        ];
                        $foundCargoCounts[$cargoKey] = true;
                    }
                }
            }
        }

        // Fallback: if no collection-based cargos found, use simple extraction
        if (empty($cargos)) {
            $seenCounts = [];

            foreach ($lines as $line) {
                if (preg_match('/(\d+)\s+(?:PALLETS?|PALLET)/i', $line, $m)) {
                    $count = (int) $m[1];

                    if (!in_array($count, $seenCounts)) {
                        $loadNumber = count($cargos) + 1;
                        $cargos[] = [
                            'title' => "Mixed Pallet Load {$loadNumber}",
                            'package_count' => $count,
                            'package_type' => 'pallet',
                            'type' => 'FTL'
                        ];
                        $seenCounts[] = $count;
                    }
                }
                elseif (preg_match('/(\d+)\s+(PACKAGES?|CARTONS?|BOXES?|ITEMS?)/i', $line, $m)) {
                    $count = (int) $m[1];

                    if (!in_array($count, $seenCounts)) {
                        $loadNumber = count($cargos) + 1;
                        $cargos[] = [
                            'title' => "Mixed Package Load {$loadNumber}",
                            'package_count' => $count,
                            'package_type' => 'package',
                            'type' => 'FTL'
                        ];
                        $seenCounts[] = $count;
                    }
                }
            }
        }

        return $cargos;
    }

    protected function extractComment(array $lines): string
    {
        $parts = [];

        foreach ($lines as $line) {
            if (preg_match('/Carrier\s+(.+?)(?:\s|$)/i', $line, $m)) {
                $parts[] = 'Carrier: ' . trim($m[1]);
            }

            if (preg_match('/Date\s+(\d{2}\/\d{2}\/\d{4})/i', $line, $m)) {
                $parts[] = 'Booking created on ' . $m[1];
            }
        }

        $hasDeliveryRestriction = collect($lines)->some(fn($line) =>
            Str::contains(Str::lower($line), 'delivery to any address other than')
        );

        if ($hasDeliveryRestriction) {
            $parts[] = 'Notes: Delivery to any address other than listed is prohibited without permission; signed POD required for payment';
        }

        return collect($parts)->filter()->implode('. ') . (count($parts) ? '.' : '');
    }

    protected function findCollectionSectionStarts(array $lines): array
    {
        $starts = [];

        for ($i = 0; $i < count($lines); $i++) {
            $line = trim($lines[$i]);
            $upperLine = strtoupper($line);
            $confidence = 0;
            $reason = '';

            // Pattern 1: Explicit "Collection COMPANY_NAME" - highest confidence
            if (preg_match('/^Collection\s+(.+)/i', $line, $m)) {
                $confidence = 100;
                $reason = 'Explicit Collection with company name';
            }

            // Pattern 2: Specific company names that are known collection points
            elseif (preg_match('/^(AKZO NOBEL|EPAC FULFILMENT SOLUTIONS LTD|LINDAL VALVE CO LTD|IBF)$/i', $line)) {
                $confidence = 95;
                $reason = 'Known collection company';
            }

            // Pattern 3: Company patterns with (C/O ...) format
            elseif (preg_match('/^([A-Z\s]+)\s*\([C\/O\s][^)]+\)$/i', $line, $m)) {
                $confidence = 90;
                $reason = 'Company with C/O pattern';
            }

            // Pattern 4: Company name with REF that appears independently
            elseif (preg_match('/^([A-Z][A-Z\s&\(\)\/\-\.C\/O]+)\s+REF\s+[A-Z0-9\/\-]+$/i', $line, $m)) {
                $nextFewLines = array_slice($lines, $i + 1, 5);
                $hasAddressInfo = false;
                foreach ($nextFewLines as $nextLine) {
                    if (preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2}|\d{5})\b/', $nextLine) ||
                        preg_match('/\b(ROAD|STREET|LANE|AVENUE|WAY|CLOSE|UNITS|HALL)\b/i', $nextLine)) {
                        $hasAddressInfo = true;
                        break;
                    }
                }
                if ($hasAddressInfo) {
                    $confidence = 85;
                    $reason = 'Company with REF + address info';
                }
            }

            // Pattern 5: Lines that look like company names followed by address patterns
            elseif (preg_match('/^([A-Z][A-Z\s&\(\)\/\-\.]{5,60})$/i', $line) &&
                     !preg_match('/\b(DELIVERY|RATE|TERMS|PAYMENT|DATE|CARRIER|ZIEGLER|ROAD|STREET|LANE|AVENUE|WAY|CLOSE|UNITS|HALL)\b/i', $line)) {
                $nextFewLines = array_slice($lines, $i + 1, 5);
                $hasAddressInfo = false;
                foreach ($nextFewLines as $nextLine) {
                    if (preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2}|\d{5})\b/', $nextLine) ||
                        preg_match('/\b(ROAD|STREET|LANE|AVENUE|WAY|CLOSE|UNITS|HALL|Sevington|Ashford)\b/i', $nextLine)) {
                        $hasAddressInfo = true;
                        break;
                    }
                }
                if ($hasAddressInfo) {
                    $confidence = 75;
                    $reason = 'Potential company name with address info';
                }
            }

            if ($confidence > 0) {
                $starts[] = [
                    'line' => $i,
                    'text' => $line,
                    'confidence' => $confidence,
                    'reason' => $reason
                ];
            }
        }

        return $starts;
    }

    protected function findCollectionStartsAggressive(array $lines): array
    {
        $starts = [];

        for ($i = 0; $i < count($lines); $i++) {
            $line = trim($lines[$i]);

            if (preg_match('/collection/i', $line)) {
                if (preg_match('/collection\s+(.+)/i', $line, $m)) {
                    $starts[] = [
                        'line' => $i,
                        'text' => $line,
                        'confidence' => 90,
                        'reason' => 'Aggressive Collection search'
                    ];
                }
            }

            elseif (preg_match('/^([A-Z][A-Z\s&\(\)\/\-\.C\/O]{8,50})\s*$/i', $line)) {
                if (isset($lines[$i + 1]) && preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2}|\d{5})\b/', $lines[$i + 1])) {
                    $starts[] = [
                        'line' => $i,
                        'text' => $line,
                        'confidence' => 60,
                        'reason' => 'Company name followed by postal code'
                    ];
                }
            }
        }

        return $starts;
    }

    protected function findCollectionSectionEnd(array $lines, int $startLine): int
    {
        $maxSectionSize = 8;

        for ($i = $startLine + 1; $i < count($lines) && $i < $startLine + $maxSectionSize; $i++) {
            $line = trim($lines[$i]);
            if (empty($line)) continue;

            if (preg_match('/^(Collection|Delivery)\s+[A-Z]/i', $line)) {
                return $i - 1;
            }

            if (preg_match('/^Rate\s*€/i', $line)) {
                return $i - 1;
            }

            if (preg_match('/^(Terms|Payment|Carrier|Date\s+\d)/i', $line)) {
                return $i - 1;
            }

            if (preg_match('/^[A-Z][A-Z\s&\(\)\/\-\.]+\s+REF\s+[A-Z0-9\/\-]+$/i', $line) && $i > $startLine + 3) {
                return $i - 1;
            }
        }

        $endLine = min(count($lines) - 1, $startLine + $maxSectionSize - 1);
        return $endLine;
    }

    protected function parseCollectionLocationSection(array $lines, int $startLine, int $endLine): ?array
    {
        $sectionLines = array_slice($lines, $startLine, $endLine - $startLine + 1);

        $result = $this->parseLocationDetails($sectionLines, 'collection');

        if ($result && isset($result['company_address'])) {
            $addr = $result['company_address'];

            if (($addr['street_address'] === 'TBC' || $addr['postal_code'] === 'TBC' || $addr['city'] === 'Unknown') &&
                $endLine < count($lines) - 1) {

                for ($i = $endLine + 1; $i <= min($endLine + 8, count($lines) - 1); $i++) {
                    if (isset($lines[$i])) {
                        $line = trim($lines[$i]);

                        if (preg_match('/^(Collection|Delivery|Rate\s*€|Terms|Payment|Carrier)/i', $line)) {
                            break;
                        }

                        $tempAddress = $addr;
                        $this->parseAddressComponents($line, $tempAddress);

                        if ($tempAddress['street_address'] !== $addr['street_address'] && $tempAddress['street_address'] !== 'TBC') {
                            $addr['street_address'] = $tempAddress['street_address'];
                        }
                        if ($tempAddress['city'] !== $addr['city'] && $tempAddress['city'] !== 'Unknown') {
                            $addr['city'] = $tempAddress['city'];
                        }
                        if ($tempAddress['postal_code'] !== $addr['postal_code'] && $tempAddress['postal_code'] !== 'TBC') {
                            $addr['postal_code'] = $tempAddress['postal_code'];
                            $addr['country_code'] = $tempAddress['country_code'];
                        }
                    }
                }

                $result['company_address'] = $addr;
            }
        }

        if ($result && isset($result['company_address'])) {
            $addr = $result['company_address'];

            if (empty($addr['company']) || $addr['company'] === 'Collection Location') {
                return null;
            }
        }

        return $result;
    }

    protected function areLocationsDuplicate(array $location1, array $location2): bool
    {
        $addr1 = $location1['company_address'];
        $addr2 = $location2['company_address'];

        if ($addr1['company'] === $addr2['company']) {
            return true;
        }

        if (isset($addr1['postal_code']) && isset($addr2['postal_code']) &&
            $addr1['postal_code'] === $addr2['postal_code']) {

            $clean1 = preg_replace('/\b(C\/O|LTD|LIMITED|CO)\b/i', '', $addr1['company']);
            $clean2 = preg_replace('/\b(C\/O|LTD|LIMITED|CO)\b/i', '', $addr2['company']);

            if (trim($clean1) === trim($clean2)) {
                return true;
            }
        }

        return false;
    }

    protected function parseLocationDetails(array $blockLines, string $type): ?array
    {
        $combinedLines = $this->combineRelatedAddressParts($blockLines);

        $company = '';
        $refComment = '';
        $address = [
            'street_address' => '',
            'city' => '',
            'postal_code' => '',
            'country_code' => 'GB'
        ];
        $timeObj = null;
        $dateFound = null;
        $timeFound = null;

        // Enhanced company extraction for multi-line format
        for ($i = 0; $i < count($combinedLines); $i++) {
            $line = trim($combinedLines[$i]);

            if (empty($company) && preg_match('/^(AKZO NOBEL|EPAC FULFILMENT SOLUTIONS LTD|LINDAL VALVE CO LTD|IBF)$/i', $line)) {
                $companyParts = [$line];

                if (isset($combinedLines[$i + 1]) && trim($combinedLines[$i + 1]) === 'REF') {
                    $i++;
                }

                if (isset($combinedLines[$i + 1]) && preg_match('/^C\/O\s+(.+)$/i', trim($combinedLines[$i + 1]), $m)) {
                    $companyParts[] = '(' . trim($combinedLines[$i + 1]) . ')';
                    $i++;
                }

                $company = implode(' ', $companyParts);
                continue;
            }
        }

        // Separate pass to extract date and time
        foreach ($combinedLines as $lineIndex => $line) {
            if (!$dateFound && preg_match('/(\d{2}\/\d{2}\/\d{4})/', $line, $m)) {
                $dateFound = $m[1];
            }

            if (!$timeFound) {
                if (preg_match('/(\d{2})(\d{2})\s*[-–]\s*(\d{1,2})\s*(pm|am)/i', $line, $m)) {
                    $startHour = (int)$m[1];
                    $startMin = (int)$m[2];
                    $endHour = (int)$m[3];
                    $endAmPm = strtolower($m[4]);

                    if ($endAmPm === 'pm' && $endHour < 12) {
                        $endHour += 12;
                    } elseif ($endAmPm === 'am' && $endHour === 12) {
                        $endHour = 0;
                    }

                    $timeFound = [
                        'start' => sprintf('%02d:%02d', $startHour, $startMin),
                        'end' => sprintf('%02d:%02d', $endHour, 0)
                    ];
                }
                elseif (preg_match('/(\d{2})(\d{2})\s*[-–]\s*(\d{2})(\d{2})/', $line, $m)) {
                    $timeFound = [
                        'start' => sprintf('%02d:%02d', (int)$m[1], (int)$m[2]),
                        'end' => sprintf('%02d:%02d', (int)$m[3], (int)$m[4])
                    ];
                }
            }
        }

        // Build time object if we have both date and time
        if ($dateFound && $timeFound) {
            try {
                $carbonDate = Carbon::createFromFormat('d/m/Y', $dateFound);
                $timeObj = [
                    'datetime_from' => $carbonDate->copy()->setTimeFromTimeString($timeFound['start'])->format('c'),
                    'datetime_to' => $carbonDate->copy()->setTimeFromTimeString($timeFound['end'])->format('c')
                ];
            } catch (\Exception $e) {
                // Continue without time object
            }
        }

        foreach ($combinedLines as $lineIndex => $line) {
            // Skip lines we've already processed in company extraction
            if (!empty($company) && (
                preg_match('/^(AKZO NOBEL|EPAC FULFILMENT SOLUTIONS LTD|LINDAL VALVE CO LTD|IBF)$/i', $line) ||
                trim($line) === 'REF' ||
                preg_match('/^C\/O\s+/i', $line)
            )) {
                continue;
            }

            // Extract company name if not already found (fallback)
            if (empty($company)) {
                $extractedCompany = $this->extractCompanyName($line);
                if ($extractedCompany) {
                    $company = $extractedCompany;
                    continue;
                }
            }

            // Extract REF comment - look for "REF XXXXX" pattern
            if (empty($refComment) && preg_match('/^REF\s+([A-Z0-9\/\-]+)$/i', $line, $m)) {
                $refComment = $m[1];
                continue;
            }
            elseif (empty($refComment) && preg_match('/(PICK\s+UP\s+[A-Z0-9]+)/i', $line, $m)) {
                $refComment = $m[1];
                continue;
            }

            // Skip lines that look like headers or section markers
            if (preg_match('/^(Collection|Delivery)\s/i', $line)) {
                continue;
            }

            // Skip standalone REF lines (but not "REF XXXXX")
            if (trim($line) === 'REF') {
                continue;
            }

            $this->parseAddressComponents($line, $address);
        }

        // Ensure we have a company name
        if (empty($company)) {
            $company = ucfirst($type) . ' Location';
        }

        // Build the final address
        $finalAddress = [
            'company' => $company,
            'street_address' => $address['street_address'] ?: 'TBC',
            'city' => $address['city'] ?: 'Unknown',
            'postal_code' => $address['postal_code'] ?: 'TBC',
            'country_code' => $address['country_code']
        ];

        // Add REF comment if found - ONLY the REF code, not company name
        if ($refComment) {
            if (preg_match('/^PICK\s+UP/i', $refComment)) {
                $finalAddress['comment'] = $refComment;
            } else {
                $finalAddress['comment'] = $refComment;
            }
        }

        $result = ['company_address' => $finalAddress];

        // Add time if found
        if ($timeObj) {
            $result['time'] = $timeObj;
        }

        return $result;
    }

    protected function combineRelatedAddressParts(array $blockLines): array
    {
        $combinedLines = [];
        $streetSuffixes = ['ROAD', 'STREET', 'LANE', 'AVENUE', 'DRIVE', 'CLOSE', 'WAY', 'GROVE', 'PARK', 'PLACE', 'CRESCENT', 'SQUARE', 'YARD', 'MEWS', 'GARDENS', 'GREEN', 'UNITS', 'HALL'];

        for ($i = 0; $i < count($blockLines); $i++) {
            $currentLine = trim($blockLines[$i]);

            $hasStreetSuffix = false;
            $currentSuffix = '';
            foreach ($streetSuffixes as $suffix) {
                if (strpos(strtoupper($currentLine), $suffix) !== false) {
                    $hasStreetSuffix = true;
                    $currentSuffix = $suffix;
                    break;
                }
            }

            if ($hasStreetSuffix) {
                $relatedParts = [$currentLine];
                $lookAheadDistance = 6;

                for ($j = $i + 1; $j <= min($i + $lookAheadDistance, count($blockLines) - 1); $j++) {
                    $nextLine = trim($blockLines[$j]);

                    // Skip cargo lines
                    if (preg_match('/^\d+\s+(PALLETS?|PACKAGES?|CARTONS?|BOXES?|ITEMS?)/i', $nextLine)) {
                        continue;
                    }

                    // Skip date/time lines
                    if (preg_match('/\d{2}\/\d{2}\/\d{4}|\d{4}-\d{4}|\d{4}-\d+[ap]m|\d{4}-\d{4}/i', $nextLine)) {
                        continue;
                    }

                    // Skip REF lines
                    if (preg_match('/^REF\s+[A-Z0-9\/\-]+$/i', $nextLine)) {
                        continue;
                    }

                    // Skip pickup instruction lines
                    if (preg_match('/^PICK\s+UP\s+[A-Z0-9]+$/i', $nextLine)) {
                        continue;
                    }

                    $nextHasStreetSuffix = false;
                    foreach ($streetSuffixes as $suffix) {
                        if (strpos(strtoupper($nextLine), $suffix) !== false) {
                            $nextHasStreetSuffix = true;
                            break;
                        }
                    }

                    if ($nextHasStreetSuffix) {
                        $relatedParts[] = $nextLine;
                        $blockLines[$j] = '';
                    }
                }

                if (count($relatedParts) > 1) {
                    $combinedAddress = implode(', ', $relatedParts);
                    $combinedLines[] = $combinedAddress;
                } else {
                    $combinedLines[] = $currentLine;
                }
            } else {
                if (!empty($currentLine)) {
                    $combinedLines[] = $currentLine;
                }
            }
        }

        return $combinedLines;
    }

    protected function extractCompanyName(string $line): string
    {
        // Pattern 1: "Collection COMPANY_NAME" or "Collection COMPANY_NAME REF XXX"
        if (preg_match('/^Collection\s+(.+?)(?:\s+REF\s+[A-Z0-9\/\-]+)?$/i', $line, $m)) {
            $company = trim($m[1]);
            return $company;
        }

        // Pattern 2: Company name with REF at end
        if (preg_match('/^(.+?)\s+REF\s+[A-Z0-9\/\-]+$/i', $line, $m)) {
            $company = trim($m[1]);
            return $company;
        }

        // Pattern 3: Company with (C/O ...) format like "AKZO NOBEL (C/O GXO LOGISTICS)"
        if (preg_match('/^([A-Z\s]+)\s*\(([C\/O\s][^)]+)\)$/i', $line, $m)) {
            $company = trim($m[1]) . ' (' . trim($m[2]) . ')';
            return $company;
        }

        // Pattern 4: Specific company patterns
        if (preg_match('/^(AKZO NOBEL|EPAC FULFILMENT SOLUTIONS LTD|LINDAL VALVE CO LTD|IBF|[A-Z][A-Z\s&\(\)\/\-\.]+(LTD|LIMITED|CO|INC|CORP|PLC|GMBH|SOLUTIONS|LOGISTICS|FULFILMENT|NOBEL|VALVE))(\s|$)/i', $line, $m)) {
            $company = trim($m[1]);
            return $company;
        }

        // Pattern 5: Two-word company names like "AKZO NOBEL"
        if (preg_match('/^([A-Z]{3,}\s+[A-Z]{3,})$/i', $line) &&
            !preg_match('/\d{2}\/\d{2}\/\d{4}/', $line) &&
            !preg_match('/\b(ROAD|STREET|LANE|AVENUE|WAY|CLOSE|UNITS|HALL)\b/i', $line)) {
            $company = trim($line);
            return $company;
        }

        // Pattern 6: All caps lines that could be company names (but not addresses or streets)
        if (preg_match('/^([A-Z][A-Z\s&\(\)\/\-\.C\/O]{5,50})$/i', $line) &&
            !preg_match('/\d{2}\/\d{2}\/\d{4}/', $line) &&
            !preg_match('/\b(ROAD|STREET|LANE|AVENUE|WAY|CLOSE|UNITS|HALL|LTD|LIMITED|SOLUTIONS|LOGISTICS|FULFILMENT)\b/i', $line) &&
            !preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2}|\d{5})\b/', $line)) {
            $company = trim($line);
            return $company;
        }

        return '';
    }

    protected function parseAddressComponents(string $line, array &$address): void
    {
        // Skip if this looks like a company name
        if (preg_match('/^(AKZO NOBEL|EPAC FULFILMENT SOLUTIONS LTD|LINDAL VALVE CO LTD|IBF)$/i', $line)) {
            return;
        }

        // Skip if this looks like a company with C/O pattern
        if (preg_match('/^C\/O\s+/i', $line)) {
            return;
        }

        // Skip Collection header lines with REF
        if (preg_match('/^Collection\s+.*REF/i', $line)) {
            return;
        }

        // Skip standalone REF line
        if (trim($line) === 'REF') {
            return;
        }

        // Skip REF with code lines
        if (preg_match('/^REF\s+[A-Z0-9\/\-]+$/i', $line)) {
            return;
        }

        // Skip if this line only contains pickup instructions
        if (preg_match('/^PICK\s+UP\s+[A-Z0-9]+$/i', $line)) {
            return;
        }

        // Skip if this line contains date/time patterns
        if (preg_match('/\d{2}\/\d{2}\/\d{4}|\d{4}-\d{4}|\d{4}-\d+[ap]m/i', $line)) {
            return;
        }

        // Skip pallet/cargo lines
        if (preg_match('/\d+\s+(pallets?|packages?)/i', $line)) {
            return;
        }

        // UK postcode + city pattern: "IP14 2QU STOWMARKET", "HAWKWELL, SS5 4JL"
        if (preg_match('/([A-Z]{1,2}\d+\s+\d[A-Z]{2})\s+([A-Z\s]+)$/i', $line, $m)) {
            $address['postal_code'] = $m[1];
            $address['city'] = trim($m[2]);
            $address['country_code'] = 'GB';
            return;
        }
        elseif (preg_match('/([A-Z\s]+),\s*([A-Z]{1,2}\d+\s+\d[A-Z]{2})$/i', $line, $m)) {
            $address['city'] = trim($m[1]);
            $address['postal_code'] = $m[2];
            $address['country_code'] = 'GB';
            return;
        }

        // French postcode + city: "95150 TAVERNY"
        if (preg_match('/^(\d{5})\s+([A-Z\s]+)$/i', $line, $m)) {
            $address['postal_code'] = $m[1];
            $address['city'] = trim($m[2]);

            if (!empty($address['postal_code'])) {
                $countryIso = GeonamesCountry::getIso($address['postal_code']);
                $address['country_code'] = $countryIso ?: 'FR';
            } else {
                $address['country_code'] = 'FR';
            }

            return;
        }

        // ENHANCED STREET ADDRESS PARSING - Check for street address first, before checking if it's already set

        // PRIORITY 1: Multi-part addresses with comma and multiple street suffixes
        if (strpos($line, ',') !== false) {
            $streetSuffixes = ['ROAD', 'STREET', 'LANE', 'AVENUE', 'DRIVE', 'CLOSE', 'WAY', 'GROVE', 'PARK', 'PLACE', 'CRESCENT', 'SQUARE', 'YARD', 'MEWS', 'GARDENS', 'GREEN', 'UNITS', 'HALL'];
            $suffixCount = 0;
            $foundSuffixes = [];
            foreach ($streetSuffixes as $suffix) {
                $count = substr_count(strtoupper($line), $suffix);
                if ($count > 0) {
                    $suffixCount += $count;
                    $foundSuffixes[] = $suffix;
                }
            }

            if ($suffixCount >= 2) {
                $address['street_address'] = trim($line);
                return;
            }
        }

        // PRIORITY 2: Specific patterns for business complex addresses like "GUSTED HALL UNITS, GUSTED HALL LANE"
        if (preg_match('/^(.+(?:UNITS|HALL),\s*.+(?:LANE|ROAD|STREET|WAY|CLOSE|CRESCENT|GROVE|PARK|PLACE).*)$/i', $line)) {
            $address['street_address'] = trim($line);
            return;
        }

        // PRIORITY 3: Other complex patterns like "CHERRYCOURT WAY, STANBRIDGE ROAD"
        if (preg_match('/^(.+(?:WAY|ROAD|STREET|LANE|AVENUE|DRIVE|CLOSE),\s*.+(?:ROAD|STREET|LANE|AVENUE|DRIVE|CLOSE).*)$/i', $line)) {
            $address['street_address'] = trim($line);
            return;
        }

        // Only proceed with single-part addresses if street_address is not already set
        if (empty($address['street_address'])) {
            // Pattern 4: Standard single street names
            if (preg_match('/^([A-Z\s]+(ROAD|RD|STREET|ST|LANE|LN|AVENUE|AVE|DRIVE|DR|CLOSE|CRESCENT|GROVE|PARK|PLACE|SQUARE|YARD|MEWS|GARDENS|GREEN|WAY))$/i', $line, $m)) {
                $address['street_address'] = trim($m[1]);
                return;
            }

            // Pattern 5: Business addresses with UNITS
            if (preg_match('/^([A-Z\s]+(UNITS|HALL|CENTRE|CENTER|INDUSTRIAL|ESTATE|BUSINESS|PARK))$/i', $line, $m)) {
                $address['street_address'] = trim($m[1]);
                return;
            }

            // Pattern 6: Simple locations
            if (preg_match('/^(Sevington|[A-Z][a-z]+)$/i', $line, $m)) {
                $address['street_address'] = trim($m[1]);
                return;
            }
        }

        // Standalone UK postcode
        if (preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2})\b/', $line, $m)) {
            if (empty($address['postal_code'])) {
                $pc = strtoupper(str_replace(' ', '', $m[1]));
                $address['postal_code'] = substr($pc, 0, -3) . ' ' . substr($pc, -3);

                if (!empty($address['postal_code'])) {
                    $countryIso = GeonamesCountry::getIso($address['postal_code']);
                    $address['country_code'] = $countryIso ?: 'GB';
                } else {
                    $address['country_code'] = 'GB';
                }
            }
        }

        // Standalone French postcode
        if (preg_match('/\b(\d{5})\b/', $line, $m)) {
            if (empty($address['postal_code'])) {
                $address['postal_code'] = $m[1];

                if (!empty($address['postal_code'])) {
                    $countryIso = GeonamesCountry::getIso($address['postal_code']);
                    $address['country_code'] = $countryIso ?: 'FR';
                } else {
                    $address['country_code'] = 'FR';
                }
            }
        }

        // City (all caps, no digits, reasonable length) - only if not already found and not a street
        if (preg_match('/^([A-Z\s\'-]{3,30}|[A-Z][a-z]+)$/i', $line) &&
            !preg_match('/\d/', $line) &&
            empty($address['city']) &&
            !preg_match('/\b(ROAD|STREET|LANE|AVENUE|WAY|CLOSE|UNITS|HALL|LTD|LIMITED|SOLUTIONS|LOGISTICS|FULFILMENT|REF|COLLECTION|DELIVERY)\b/i', $line)) {
            $address['city'] = trim($line);
        }
    }
}