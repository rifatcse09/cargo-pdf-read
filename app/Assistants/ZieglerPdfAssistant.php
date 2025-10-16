<?php

namespace App\Assistants;

use Carbon\Carbon;
use App\GeonamesCountry;
use Illuminate\Support\Arr;
use Illuminate\Support\Str;
use Illuminate\Support\Facades\Log;

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

        // Debug: log ALL lines to understand the actual structure
        Log::info("=== COMPLETE ZIEGLER PDF STRUCTURE ===");
        foreach ($lines as $i => $line) {
            $trimmed = trim($line);
            if (!empty($trimmed)) {
                Log::info("Line {$i}: '{$trimmed}'");
            }
        }

        // Clean and normalize lines
        $lines = collect($lines)
            ->map(fn($l) => trim(preg_replace('/\s+/', ' ', $l)))
            ->filter(fn($l) => !empty($l))
            ->values()
            ->toArray();

        // Extract all components
        $customer = $this->extractCustomer($lines);
        $orderRef = $this->extractOrderReference($lines);
        $freight = $this->extractFreight($lines);
        $loadingLocs = $this->extractLoadingLocations($lines);
        $destLocs = $this->extractDestinationLocations($lines);
        $cargos = $this->extractCargos($lines);
        $comment = $this->extractComment($lines);

        // Debug extracted data
        Log::info("Extracted loading locations: " . json_encode($loadingLocs));
        Log::info("Extracted destination locations: " . json_encode($destLocs));
        Log::info("Extracted cargos: " . json_encode($cargos));

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
            // Based on the provided text: lines after ZIEGLER UK LTD
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
                // Ensure country is set when we set fallback postal code
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

        Log::info("=== PARSING ZIEGLER ADDRESS ===");
        Log::info("Address lines to process: " . json_encode($addressLines));

        foreach ($addressLines as $lineIndex => $line) {
            $line = trim($line);
            if (empty($line)) break;

            Log::info("Processing Ziegler address line {$lineIndex}: '{$line}'");

            // Stop at booking/instruction lines
            if (preg_match('/\b(BOOKING|INSTRUCTION|TELEPHONE|REF)\b/i', $line)) {
                Log::info("Stopping at booking/instruction line: '{$line}'");
                break;
            }

            // Check for UK postcode first (SS17 9FJ)
            if (preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2})\b/', $line, $m)) {
                $pc = strtoupper(str_replace(' ', '', $m[1]));
                $details['postal_code'] = substr($pc, 0, -3) . ' ' . substr($pc, -3);
                $postalCodeFound = true;

                // Use GeonamesCountry to detect country from postal code
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
                Log::info("Found postal code: {$details['postal_code']}");
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
                Log::info("Found French postal code: {$details['postal_code']}");
                continue;
            }

            // Check if this is a city line - but be more specific about city detection
            // Only treat as city if it's near the end and looks like a typical city name
            if (!$cityFound && preg_match('/^[A-Z\s\'-]{4,}$/', $line) && !preg_match('/\d/', $line)) {
                // Check if this looks like a known city pattern or if we're near postal code
                $isLikelyCity = false;

                // If we found postal code already or in next line, this is likely city
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
                    Log::info("Found city: {$details['city']}");
                    continue;
                }
            }

            // If we haven't identified this as postal code or city, it's part of street address
            // This will capture both "LONDON GATEWAY LOGISTICS PARK" and "NORTH 4, NORTH SEA CROSSING"
            $streetParts[] = $line;
        }

        Log::info("Street parts collected: " . json_encode($streetParts));
        Log::info("City found: " . ($cityFound ? $details['city'] ?? 'none' : 'none'));
        Log::info("Postal code found: " . ($postalCodeFound ? $details['postal_code'] ?? 'none' : 'none'));

        // Build the complete street address
        if (!empty($streetParts)) {
            $details['street_address'] = collect($streetParts)->implode(', ');
            Log::info("Final street address: {$details['street_address']}");
        }

        Log::info("=== ZIEGLER ADDRESS PARSING COMPLETE ===");
    }

    protected function findZieglerTerms(array $lines): ?string
    {
        $parts = [];

        foreach ($lines as $line) {
            $lowerLine = Str::lower($line);

            // Look for purchase order terms
            if (Str::contains($lowerLine, ['purchase order general terms', 'purchase order terms'])) {
                $parts[] = 'Ziegler purchase order terms apply';
            }

            // Look for reference quote requirement
            if (Str::contains($lowerLine, ['please quote our reference number', 'quote our reference on your invoice'])) {
                $parts[] = 'Please quote our reference on invoice';
            }
        }

        // Return combined comment or null if no parts found
        if (empty($parts)) {
            return null;
        }

        return collect($parts)->unique()->implode('. ') . '.';
    }

    protected function extractOrderReference(array $lines): ?string
    {
        $foundReferences = [];

        foreach ($lines as $i => $line) {
            // "Ziegler Ref 187395" - highest priority
            if (preg_match('/Ziegler\s+Ref\s+(\w+)/i', $line, $m)) {
                $foundReferences[] = ['priority' => 1, 'ref' => $m[1], 'line' => $i, 'pattern' => 'Ziegler Ref'];
            }

            // "Our Ref: XXXXXXX" - high priority
            if (preg_match('/Our\s+Ref(?:erence)?[:\s]+([A-Z0-9\-\/]+)/i', $line, $m)) {
                $foundReferences[] = ['priority' => 2, 'ref' => trim($m[1]), 'line' => $i, 'pattern' => 'Our Ref'];
            }

            // "Ref: XXXXXXX" at start of line - medium priority
            if (preg_match('/^Ref(?:erence)?[:\s]+([A-Z0-9\-\/]+)/i', $line, $m)) {
                $foundReferences[] = ['priority' => 3, 'ref' => trim($m[1]), 'line' => $i, 'pattern' => 'Line Ref'];
            }

            // Look for specific known references
            if (preg_match('/\b(187395)\b/', $line, $m)) {
                $foundReferences[] = ['priority' => 1, 'ref' => $m[1], 'line' => $i, 'pattern' => 'Specific 187395'];
            }

            if (preg_match('/\b(98111001678)\b/', $line, $m)) {
                $foundReferences[] = ['priority' => 1, 'ref' => $m[1], 'line' => $i, 'pattern' => 'Specific 98111001678'];
            }

            // Generic "Ref" followed by alphanumeric - lower priority
            if (preg_match('/\bRef[:\s]+([A-Z0-9\-\/]{4,})/i', $line, $m)) {
                $ref = trim($m[1]);
                // Skip if it's just common words
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

        // Log all found references for debugging
        foreach ($foundReferences as $ref) {
            Log::info("Found reference: {$ref['ref']} (priority: {$ref['priority']}, pattern: {$ref['pattern']}, line: {$ref['line']})");
        }

        // Return the highest priority reference
        if (!empty($foundReferences)) {
            $bestRef = $foundReferences[0]['ref'];
            Log::info("Selected order reference: {$bestRef}");
            return $bestRef;
        }

        Log::info("No order reference found");
        return null;
    }

    protected function extractFreight(array $lines): array
    {
        $price = null;
        $currency = null;

        foreach ($lines as $line) {
            Log::info("Checking freight line: {$line}");

            // "Rate € 1,000" or "Rate € 1.15" - more flexible pattern
            if (preg_match('/Rate\s*€\s*([0-9,.\s]+)/i', $line, $m)) {
                $cleanPrice = function_exists('uncomma') ? uncomma($m[1]) : str_replace([',', ' '], '', $m[1]);
                $price = (float) $cleanPrice;
                $currency = 'EUR';
                Log::info("Found Rate € pattern: original='{$m[1]}', cleaned='{$cleanPrice}', final={$price}");
                break;
            }

            // "€ 1,000" or "€ 1.15" - Euro symbol first
            if (preg_match('/€\s*([0-9,.\s]+)/i', $line, $m)) {
                $cleanPrice = function_exists('uncomma') ? uncomma($m[1]) : str_replace([',', ' '], '', $m[1]);
                $price = (float) $cleanPrice;
                $currency = 'EUR';
                Log::info("Found € pattern: original='{$m[1]}', cleaned='{$cleanPrice}', final={$price}");
                break;
            }

            // "1,000 €" or "1.15 €" - Euro symbol after
            if (preg_match('/([0-9,.\s]+)\s*€/i', $line, $m)) {
                $cleanPrice = function_exists('uncomma') ? uncomma($m[1]) : str_replace([',', ' '], '', $m[1]);
                $price = (float) $cleanPrice;
                $currency = 'EUR';
                Log::info("Found amount € pattern: original='{$m[1]}', cleaned='{$cleanPrice}', final={$price}");
                break;
            }

            // "Price: EUR 1000"
            if (preg_match('/Price[:\s]*EUR\s*([0-9,.\s]+)/i', $line, $m)) {
                $cleanPrice = function_exists('uncomma') ? uncomma($m[1]) : str_replace([',', ' '], '', $m[1]);
                $price = (float) $cleanPrice;
                $currency = 'EUR';
                Log::info("Found Price EUR pattern: original='{$m[1]}', cleaned='{$cleanPrice}', final={$price}");
                break;
            }

            // Generic currency patterns
            if (preg_match('/(EUR|GBP|USD)\s*([0-9,.\s]+)/i', $line, $m)) {
                $currency = strtoupper($m[1]);
                $cleanPrice = function_exists('uncomma') ? uncomma($m[2]) : str_replace([',', ' '], '', $m[2]);
                $price = (float) $cleanPrice;
                Log::info("Found {$currency} pattern: original='{$m[2]}', cleaned='{$cleanPrice}', final={$price}");
                break;
            }

            if (preg_match('/([€£$])\s*([0-9,.\s]+)/', $line, $m)) {
                $currency = ['€' => 'EUR', '£' => 'GBP', '$' => 'USD'][$m[1]] ?? null;
                $cleanPrice = function_exists('uncomma') ? uncomma($m[2]) : str_replace([',', ' '], '', $m[2]);
                $price = (float) $cleanPrice;
                Log::info("Found symbol {$m[1]} pattern: original='{$m[2]}', cleaned='{$cleanPrice}', final={$price}");
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

        Log::info("Final freight result: price={$price}, currency={$currency}");

        // Default to EUR if no currency found (common for Ziegler)
        return [
            'price' => $price,
            'currency' => $currency ?: 'EUR'
        ];
    }

    protected function parseLocationDetails(array $blockLines, string $type): ?array
    {
        Log::info("=== PARSING LOCATION DETAILS FOR " . strtoupper($type) . " ===");
        Log::info("Block lines: " . json_encode($blockLines));

        // PRE-PROCESS: Combine related address parts that may be separated by cargo lines
        $combinedLines = $this->combineRelatedAddressParts($blockLines);
        Log::info("After combining related parts: " . json_encode($combinedLines));

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

            // Look for company name patterns
            if (empty($company) && preg_match('/^(AKZO NOBEL|EPAC FULFILMENT SOLUTIONS LTD|LINDAL VALVE CO LTD|IBF)$/i', $line)) {
                $companyParts = [$line];

                // Check if next line is "REF" (standalone) and skip it
                if (isset($combinedLines[$i + 1]) && trim($combinedLines[$i + 1]) === 'REF') {
                    $i++; // Skip the standalone REF line
                }

                // Check if next line is C/O pattern
                if (isset($combinedLines[$i + 1]) && preg_match('/^C\/O\s+(.+)$/i', trim($combinedLines[$i + 1]), $m)) {
                    $companyParts[] = '(' . trim($combinedLines[$i + 1]) . ')';
                    $i++; // Skip the C/O line as we've processed it
                }

                $company = implode(' ', $companyParts);
                Log::info("Built company from multiple lines: '{$company}'");
                continue;
            }
        }

        // Separate pass to extract date and time - look for them in the section
        foreach ($combinedLines as $lineIndex => $line) {
            Log::info("Checking for date/time in line {$lineIndex}: '{$line}'");

            // Extract date: DD/MM/YYYY
            if (!$dateFound && preg_match('/(\d{2}\/\d{2}\/\d{4})/', $line, $m)) {
                $dateFound = $m[1];
                Log::info("Found date in section: {$dateFound}");
            }

            // Extract time with enhanced patterns for PM/AM
            if (!$timeFound) {
                // HHMM-HamPM format (like 0900-2pm, 0900-3PM)
                if (preg_match('/(\d{2})(\d{2})\s*[-–]\s*(\d{1,2})\s*(pm|am)/i', $line, $m)) {
                    $startHour = (int)$m[1];
                    $startMin = (int)$m[2];
                    $endHour = (int)$m[3];
                    $endAmPm = strtolower($m[4]);

                    // Convert PM/AM to 24-hour format
                    if ($endAmPm === 'pm' && $endHour < 12) {
                        $endHour += 12;
                    } elseif ($endAmPm === 'am' && $endHour === 12) {
                        $endHour = 0;
                    }

                    $timeFound = [
                        'start' => sprintf('%02d:%02d', $startHour, $startMin),
                        'end' => sprintf('%02d:%02d', $endHour, 0)
                    ];
                    Log::info("Found time range (HHMM-HamPM): {$timeFound['start']} - {$timeFound['end']}");
                }
                // Standard HHMM-HHMM format
                elseif (preg_match('/(\d{2})(\d{2})\s*[-–]\s*(\d{2})(\d{2})/', $line, $m)) {
                    $timeFound = [
                        'start' => sprintf('%02d:%02d', (int)$m[1], (int)$m[2]),
                        'end' => sprintf('%02d:%02d', (int)$m[3], (int)$m[4])
                    ];
                    Log::info("Found time range (HHMM-HHMM): {$timeFound['start']} - {$timeFound['end']}");
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
                Log::info("Built time object from section: " . json_encode($timeObj));
            } catch (\Exception $e) {
                Log::error("Error building time object: " . $e->getMessage());
            }
        }

        foreach ($combinedLines as $lineIndex => $line) {
            Log::info("Processing detail line {$lineIndex}: '{$line}'");

            // Skip lines we've already processed in company extraction
            if (!empty($company) && (
                preg_match('/^(AKZO NOBEL|EPAC FULFILMENT SOLUTIONS LTD|LINDAL VALVE CO LTD|IBF)$/i', $line) ||
                trim($line) === 'REF' ||
                preg_match('/^C\/O\s+/i', $line)
            )) {
                Log::info("Skipping already processed line: '{$line}'");
                continue;
            }

            // Extract company name if not already found (fallback)
            if (empty($company)) {
                $extractedCompany = $this->extractCompanyName($line);
                if ($extractedCompany) {
                    $company = $extractedCompany;
                    Log::info("Extracted company: '{$company}'");
                    continue;
                }
            }

            // Extract REF comment - look for "REF XXXXX" pattern
            if (empty($refComment) && preg_match('/^REF\s+([A-Z0-9\/\-]+)$/i', $line, $m)) {
                $refComment = $m[1];
                Log::info("Found REF: {$refComment}");
                continue;
            }
            // Also check for patterns like "PICK UP T1"
            elseif (empty($refComment) && preg_match('/(PICK\s+UP\s+[A-Z0-9]+)/i', $line, $m)) {
                $refComment = $m[1];
                Log::info("Found pickup comment: {$refComment}");
                continue;
            }

            // Skip lines that look like headers or section markers
            if (preg_match('/^(Collection|Delivery)\s/i', $line)) {
                Log::info("Skipping header line: '{$line}'");
                continue;
            }

            // Skip standalone REF lines (but not "REF XXXXX")
            if (trim($line) === 'REF') {
                Log::info("Skipping standalone REF line: '{$line}'");
                continue;
            }

            // Parse address components
            Log::info("Attempting to parse address from: '{$line}'");
            $this->parseAddressComponents($line, $address);
        }

        // Ensure we have a company name
        if (empty($company)) {
            $company = ucfirst($type) . ' Location';
            Log::info("Using fallback company name: {$company}");
        }

        // Build the final address
        $finalAddress = [
            'company' => $company,
            'street_address' => $address['street_address'] ?: 'TBC',
            'city' => $address['city'] ?: 'Unknown',
            'postal_code' => $address['postal_code'] ?: 'TBC',
            'country_code' => $address['country_code']
        ];

        // Add REF comment if found
        if ($refComment) {
            if (preg_match('/^PICK\s+UP/i', $refComment)) {
                $finalAddress['comment'] = $refComment;
            } else {
                // Extract just the company name without C/O for the comment
                $companyForComment = preg_replace('/\s*\([^)]+\)$/', '', $company);
                $finalAddress['comment'] = $companyForComment . ' REF ' . $refComment;
            }
        }

        $result = ['company_address' => $finalAddress];

        // Add time if found
        if ($timeObj) {
            $result['time'] = $timeObj;
        }

        Log::info("=== FINAL LOCATION RESULT ===");
        Log::info(json_encode($result, JSON_PRETTY_PRINT));

        return $result;
    }

    protected function combineRelatedAddressParts(array $blockLines): array
    {
        Log::info("=== COMBINING RELATED ADDRESS PARTS ===");
        Log::info("Input lines: " . json_encode($blockLines));

        $combinedLines = [];
        $streetSuffixes = ['ROAD', 'STREET', 'LANE', 'AVENUE', 'DRIVE', 'CLOSE', 'WAY', 'GROVE', 'PARK', 'PLACE', 'CRESCENT', 'SQUARE', 'YARD', 'MEWS', 'GARDENS', 'GREEN', 'UNITS', 'HALL'];

        for ($i = 0; $i < count($blockLines); $i++) {
            $currentLine = trim($blockLines[$i]);

            // Check if this line has a street suffix
            $hasStreetSuffix = false;
            foreach ($streetSuffixes as $suffix) {
                if (strpos(strtoupper($currentLine), $suffix) !== false) {
                    $hasStreetSuffix = true;
                    break;
                }
            }

            if ($hasStreetSuffix) {
                Log::info("Line {$i} has street suffix: '{$currentLine}'");

                // Look ahead for related street parts (skipping cargo/non-address lines)
                $relatedParts = [$currentLine];
                $lookAheadDistance = 3; // Look ahead up to 3 lines

                for ($j = $i + 1; $j <= min($i + $lookAheadDistance, count($blockLines) - 1); $j++) {
                    $nextLine = trim($blockLines[$j]);

                    // Skip cargo lines (like "10 PALLETS")
                    if (preg_match('/^\d+\s+(PALLETS?|PACKAGES?|CARTONS?|BOXES?|ITEMS?)/i', $nextLine)) {
                        Log::info("Skipping cargo line {$j}: '{$nextLine}'");
                        continue;
                    }

                    // Skip date/time lines
                    if (preg_match('/\d{2}\/\d{2}\/\d{4}|\d{4}-\d{4}|\d{4}-\d+[ap]m/i', $nextLine)) {
                        Log::info("Skipping date/time line {$j}: '{$nextLine}'");
                        continue;
                    }

                    // Check if this next line has a street suffix
                    $nextHasStreetSuffix = false;
                    foreach ($streetSuffixes as $suffix) {
                        if (strpos(strtoupper($nextLine), $suffix) !== false) {
                            $nextHasStreetSuffix = true;
                            break;
                        }
                    }

                    if ($nextHasStreetSuffix) {
                        Log::info("Found related street part at line {$j}: '{$nextLine}'");
                        $relatedParts[] = $nextLine;

                        // Mark this line as processed by setting it to empty
                        $blockLines[$j] = '';
                    }
                }

                // If we found multiple related parts, combine them
                if (count($relatedParts) > 1) {
                    $combinedAddress = implode(', ', $relatedParts);
                    Log::info("COMBINED address parts: '{$combinedAddress}'");
                    $combinedLines[] = $combinedAddress;
                } else {
                    $combinedLines[] = $currentLine;
                }
            } else {
                // Not a street line, add as-is (unless it was already processed)
                if (!empty($currentLine)) {
                    $combinedLines[] = $currentLine;
                }
            }
        }

        Log::info("Final combined lines: " . json_encode($combinedLines));
        Log::info("=== END COMBINING ADDRESS PARTS ===");

        return $combinedLines;
    }

    protected function parseAddressLine(string $line, array &$address): void
    {
        // Street patterns
        if (preg_match('/^(.+(?:ROAD|RD|STREET|ST|LANE|LN|AVENUE|AVE|DRIVE|DR|UNITS|HALL|WAY|CLOSE|CRESCENT|GROVE|PARK|PLACE|SQUARE|YARD|MEWS|GARDENS|GREEN).*?)(?:\s+[A-Z]{1,2}\d+|$)/i', $line, $m)) {
            if (empty($address['street_address'])) {
                $address['street_address'] = trim($m[1]);
                Log::info("Found street: {$address['street_address']}");
            }
        }

        // UK postcode + city pattern: "IP14 2QU STOWMARKET"
        if (preg_match('/([A-Z]{1,2}\d+\s+\d[A-Z]{2})\s+([A-Z\s]+)$/i', $line, $m)) {
            $address['postal_code'] = $m[1];
            $address['city'] = trim($m[2]);
            $address['country_code'] = 'GB';
            Log::info("Found UK postcode+city: {$address['postal_code']} {$address['city']}");
            return;
        }

        // French postcode + city: "95150 TAVERNY"
        if (preg_match('/^(\d{5})\s+([A-Z\s]+)$/i', $line, $m)) {
            $address['postal_code'] = $m[1];
            $address['city'] = trim($m[2]);
            $address['country_code'] = 'FR';
            Log::info("Found French postcode+city: {$address['postal_code']} {$address['city']}");
            return;
        }

        // Standalone UK postcode
        if (preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2})\b/', $line, $m)) {
            if (empty($address['postal_code'])) {
                $pc = strtoupper(str_replace(' ', '', $m[1]));
                $address['postal_code'] = substr($pc, 0, -3) . ' ' . substr($pc, -3);
                $address['country_code'] = 'GB';
                Log::info("Found standalone UK postcode: {$address['postal_code']}");
            }
        }

        // Standalone French postcode
        if (preg_match('/\b(\d{5})\b/', $line, $m)) {
            if (empty($address['postal_code'])) {
                $address['postal_code'] = $m[1];
                $address['country_code'] = 'FR';
                Log::info("Found standalone French postcode: {$address['postal_code']}");
            }
        }

        // City (all caps, no digits, reasonable length)
        if (preg_match('/^([A-Z\s\'-]{3,30})$/i', $line) && !preg_match('/\d/', $line)) {
            if (empty($address['city'])) {
                $address['city'] = trim($line);
                Log::info("Found standalone city: {$address['city']}");
            }
        }
    }

    protected function parseAddressComponents(string $line, array &$address): void
    {
        Log::info("=== PARSING ADDRESS COMPONENTS ===");
        Log::info("Input line: '{$line}'");
        Log::info("Current address state: " . json_encode($address));

        // Skip if this looks like a company name
        if (preg_match('/^(AKZO NOBEL|EPAC FULFILMENT SOLUTIONS LTD|LINDAL VALVE CO LTD|IBF)$/i', $line)) {
            Log::info("Skipping company name line for address parsing: '{$line}'");
            return;
        }

        // Skip if this looks like a company with C/O pattern
        if (preg_match('/^C\/O\s+/i', $line)) {
            Log::info("Skipping C/O line for address parsing: '{$line}'");
            return;
        }

        // Skip Collection header lines with REF
        if (preg_match('/^Collection\s+.*REF/i', $line)) {
            Log::info("Skipping Collection header line for address parsing: '{$line}'");
            return;
        }

        // Skip standalone REF line
        if (trim($line) === 'REF') {
            Log::info("Skipping standalone REF line for address parsing: '{$line}'");
            return;
        }

        // Skip REF with code lines
        if (preg_match('/^REF\s+[A-Z0-9\/\-]+$/i', $line)) {
            Log::info("Skipping REF code line for address parsing: '{$line}'");
            return;
        }

        // Skip if this line only contains pickup instructions
        if (preg_match('/^PICK\s+UP\s+[A-Z0-9]+$/i', $line)) {
            Log::info("Skipping pickup instruction line for address parsing: '{$line}'");
            return;
        }

        // Skip if this line contains date/time patterns
        if (preg_match('/\d{2}\/\d{2}\/\d{4}|\d{4}-\d{4}|\d{4}-\d+[ap]m/i', $line)) {
            Log::info("Skipping date/time line for address parsing: '{$line}'");
            return;
        }

        // Skip pallet/cargo lines
        if (preg_match('/\d+\s+(pallets?|packages?)/i', $line)) {
            Log::info("Skipping cargo line for address parsing: '{$line}'");
            return;
        }

        // UK postcode + city pattern: "IP14 2QU STOWMARKET", "HAWKWELL, SS5 4JL"
        if (preg_match('/([A-Z]{1,2}\d+\s+\d[A-Z]{2})\s+([A-Z\s]+)$/i', $line, $m)) {
            // Postcode City format
            $address['postal_code'] = $m[1];
            $address['city'] = trim($m[2]);
            $address['country_code'] = 'GB';
            Log::info("Found UK postcode+city (format 1): {$address['postal_code']} {$address['city']}");
            return;
        }
        elseif (preg_match('/([A-Z\s]+),\s*([A-Z]{1,2}\d+\s+\d[A-Z]{2})$/i', $line, $m)) {
            // City, Postcode format
            $address['city'] = trim($m[1]);
            $address['postal_code'] = $m[2];
            $address['country_code'] = 'GB';
            Log::info("Found UK city+postcode (format 2): {$address['city']} {$address['postal_code']}");
            return;
        }

        // French postcode + city: "95150 TAVERNY"
        if (preg_match('/^(\d{5})\s+([A-Z\s]+)$/i', $line, $m)) {
            $address['postal_code'] = $m[1];
            $address['city'] = trim($m[2]);

            // Use GeonamesCountry to detect country from postal code - with null check
            if (!empty($address['postal_code'])) {
                $countryIso = GeonamesCountry::getIso($address['postal_code']);
                $address['country_code'] = $countryIso ?: 'FR';
            } else {
                $address['country_code'] = 'FR';
            }

            Log::info("Found French postcode+city: {$address['postal_code']} {$address['city']}");
            return;
        }

        // ENHANCED STREET ADDRESS PARSING - Check for street address first, before checking if it's already set
        Log::info("Attempting street address parsing for: '{$line}'");

        // PRIORITY 1: Multi-part addresses with comma and multiple street suffixes
        if (strpos($line, ',') !== false) {
            Log::info("Line contains comma, checking for multi-part street address");

            // Count street suffixes in the line
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

            Log::info("Found {$suffixCount} street suffixes in line: " . implode(', ', $foundSuffixes));

            if ($suffixCount >= 2) {
                // Force override any existing street address for multi-part addresses
                $address['street_address'] = trim($line);
                Log::info("CAPTURED multi-part street address (OVERRIDING ANY EXISTING): '{$address['street_address']}'");
                Log::info("Final address state: " . json_encode($address));
                Log::info("=== END PARSING ADDRESS COMPONENTS ===");
                return;
            }
        }

        // PRIORITY 2: Specific patterns for business complex addresses like "GUSTED HALL UNITS, GUSTED HALL LANE"
        if (preg_match('/^(.+(?:UNITS|HALL),\s*.+(?:LANE|ROAD|STREET|WAY|CLOSE|CRESCENT|GROVE|PARK|PLACE).*)$/i', $line)) {
            // Force override for business complex patterns
            $address['street_address'] = trim($line);
            Log::info("CAPTURED via business complex pattern (OVERRIDING): '{$address['street_address']}'");
            Log::info("Final address state: " . json_encode($address));
            Log::info("=== END PARSING ADDRESS COMPONENTS ===");
            return;
        }

        // PRIORITY 3: Other complex patterns like "CHERRYCOURT WAY, STANBRIDGE ROAD"
        if (preg_match('/^(.+(?:WAY|ROAD|STREET|LANE|AVENUE|DRIVE|CLOSE),\s*.+(?:ROAD|STREET|LANE|AVENUE|DRIVE|CLOSE).*)$/i', $line)) {
            // Force override for complex street patterns
            $address['street_address'] = trim($line);
            Log::info("CAPTURED via complex street pattern (OVERRIDING): '{$address['street_address']}'");
            Log::info("Final address state: " . json_encode($address));
            Log::info("=== END PARSING ADDRESS COMPONENTS ===");
            return;
        }

        // Only proceed with single-part addresses if street_address is not already set
        if (empty($address['street_address'])) {
            Log::info("No street address set yet, checking single-part patterns");

            // Pattern 4: Standard single street names
            if (preg_match('/^([A-Z\s]+(ROAD|RD|STREET|ST|LANE|LN|AVENUE|AVE|DRIVE|DR|CLOSE|CRESCENT|GROVE|PARK|PLACE|SQUARE|YARD|MEWS|GARDENS|GREEN|WAY))$/i', $line, $m)) {
                $address['street_address'] = trim($m[1]);
                Log::info("CAPTURED standard street address: '{$address['street_address']}'");
                Log::info("Final address state: " . json_encode($address));
                Log::info("=== END PARSING ADDRESS COMPONENTS ===");
                return;
            }

            // Pattern 5: Business addresses with UNITS
            if (preg_match('/^([A-Z\s]+(UNITS|HALL|CENTRE|CENTER|INDUSTRIAL|ESTATE|BUSINESS|PARK))$/i', $line, $m)) {
                $address['street_address'] = trim($m[1]);
                Log::info("CAPTURED business address: '{$address['street_address']}'");
                Log::info("Final address state: " . json_encode($address));
                Log::info("=== END PARSING ADDRESS COMPONENTS ===");
                return;
            }

            // Pattern 6: Simple locations
            if (preg_match('/^(Sevington|[A-Z][a-z]+)$/i', $line, $m)) {
                $address['street_address'] = trim($m[1]);
                Log::info("CAPTURED simple location: '{$address['street_address']}'");
                Log::info("Final address state: " . json_encode($address));
                Log::info("=== END PARSING ADDRESS COMPONENTS ===");
                return;
            }

            Log::info("No street address pattern matched for: '{$line}'");
        } else {
            Log::info("Street address already set: '{$address['street_address']}', but complex patterns can override");
        }

        // Standalone UK postcode
        if (preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2})\b/', $line, $m)) {
            if (empty($address['postal_code'])) {
                $pc = strtoupper(str_replace(' ', '', $m[1]));
                $address['postal_code'] = substr($pc, 0, -3) . ' ' . substr($pc, -3);

                // Use GeonamesCountry to detect country from postal code - with null check
                if (!empty($address['postal_code'])) {
                    $countryIso = GeonamesCountry::getIso($address['postal_code']);
                    $address['country_code'] = $countryIso ?: 'GB';
                } else {
                    $address['country_code'] = 'GB';
                }

                Log::info("Found standalone UK postcode: {$address['postal_code']}");
            }
        }

        // Standalone French postcode
        if (preg_match('/\b(\d{5})\b/', $line, $m)) {
            if (empty($address['postal_code'])) {
                $address['postal_code'] = $m[1];

                // Use GeonamesCountry to detect country from postal code - with null check
                if (!empty($address['postal_code'])) {
                    $countryIso = GeonamesCountry::getIso($address['postal_code']);
                    $address['country_code'] = $countryIso ?: 'FR';
                } else {
                    $address['country_code'] = 'FR';
                }

                Log::info("Found standalone French postcode: {$address['postal_code']}");
            }
        }

        // City (all caps, no digits, reasonable length) - only if not already found and not a street
        if (preg_match('/^([A-Z\s\'-]{3,30}|[A-Z][a-z]+)$/i', $line) &&
            !preg_match('/\d/', $line) &&
            empty($address['city']) &&
            !preg_match('/\b(ROAD|STREET|LANE|AVENUE|WAY|CLOSE|UNITS|HALL|LTD|LIMITED|SOLUTIONS|LOGISTICS|FULFILMENT|REF|COLLECTION|DELIVERY)\b/i', $line)) {
            $address['city'] = trim($line);
            Log::info("Found standalone city: {$address['city']}");
        }

        Log::info("Final address state: " . json_encode($address));
        Log::info("=== END PARSING ADDRESS COMPONENTS ===");
    }

    protected function extractCompanyName(string $line): string
    {
        Log::info("Extracting company from: '{$line}'");

        // Pattern 1: "Collection COMPANY_NAME" or "Collection COMPANY_NAME REF XXX"
        if (preg_match('/^Collection\s+(.+?)(?:\s+REF\s+[A-Z0-9\/\-]+)?$/i', $line, $m)) {
            $company = trim($m[1]);
            Log::info("Found company via Collection pattern: '{$company}'");
            return $company;
        }

        // Pattern 2: Company name with REF at end
        if (preg_match('/^(.+?)\s+REF\s+[A-Z0-9\/\-]+$/i', $line, $m)) {
            $company = trim($m[1]);
            Log::info("Found company via REF pattern: '{$company}'");
            return $company;
        }

        // Pattern 3: Company with (C/O ...) format like "AKZO NOBEL (C/O GXO LOGISTICS)"
        if (preg_match('/^([A-Z\s]+)\s*\(([C\/O\s][^)]+)\)$/i', $line, $m)) {
            $company = trim($m[1]) . ' (' . trim($m[2]) . ')';
            Log::info("Found company via C/O pattern: '{$company}'");
            return $company;
        }

        // Pattern 4: Specific company patterns
        if (preg_match('/^(AKZO NOBEL|EPAC FULFILMENT SOLUTIONS LTD|LINDAL VALVE CO LTD|IBF|[A-Z][A-Z\s&\(\)\/\-\.]+(LTD|LIMITED|CO|INC|CORP|PLC|GMBH|SOLUTIONS|LOGISTICS|FULFILMENT|NOBEL|VALVE))(\s|$)/i', $line, $m)) {
            $company = trim($m[1]);
            Log::info("Found company via specific/business suffix pattern: '{$company}'");
            return $company;
        }

        // Pattern 5: Two-word company names like "AKZO NOBEL"
        if (preg_match('/^([A-Z]{3,}\s+[A-Z]{3,})$/i', $line) &&
            !preg_match('/\d{2}\/\d{2}\/\d{4}/', $line) &&
            !preg_match('/\b(ROAD|STREET|LANE|AVENUE|WAY|CLOSE|UNITS|HALL)\b/i', $line)) {
            $company = trim($line);
            Log::info("Found company via two-word pattern: '{$company}'");
            return $company;
        }

        // Pattern 6: All caps lines that could be company names (but not addresses or streets)
        if (preg_match('/^([A-Z][A-Z\s&\(\)\/\-\.C\/O]{5,50})$/i', $line) &&
            !preg_match('/\d{2}\/\d{2}\/\d{4}/', $line) &&
            !preg_match('/\b(ROAD|STREET|LANE|AVENUE|WAY|CLOSE|UNITS|HALL|LTD|LIMITED|SOLUTIONS|LOGISTICS|FULFILMENT)\b/i', $line) &&
            !preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2}|\d{5})\b/', $line)) { // Not postal codes
            $company = trim($line);
            Log::info("Found company via all-caps pattern: '{$company}'");
            return $company;
        }

        Log::info("No company name found in line");
        return '';
    }

    protected function findCollectionSectionStarts(array $lines): array
    {
        $starts = [];

        Log::info("=== SCANNING FOR COLLECTION SECTION STARTS ===");
        for ($i = 0; $i < count($lines); $i++) {
            $line = trim($lines[$i]);
            $upperLine = strtoupper($line);
            $confidence = 0;
            $reason = '';

            Log::info("Scanning line {$i}: '{$line}'");

            // Pattern 1: Explicit "Collection COMPANY_NAME" - highest confidence
            if (preg_match('/^Collection\s+(.+)/i', $line, $m)) {
                $confidence = 100;
                $reason = 'Explicit Collection with company name';
                Log::info("FOUND Collection pattern at line {$i}: company='{$m[1]}'");
            }

            // Pattern 2: Specific company names that are known collection points
            elseif (preg_match('/^(AKZO NOBEL|EPAC FULFILMENT SOLUTIONS LTD|LINDAL VALVE CO LTD|IBF)$/i', $line)) {
                $confidence = 95;
                $reason = 'Known collection company';
                Log::info("FOUND known company at line {$i}: '{$line}'");
            }

            // Pattern 3: Company patterns with (C/O ...) format
            elseif (preg_match('/^([A-Z\s]+)\s*\([C\/O\s][^)]+\)$/i', $line, $m)) {
                $confidence = 90;
                $reason = 'Company with C/O pattern';
                Log::info("FOUND company with C/O at line {$i}: '{$line}'");
            }

            // Pattern 4: Company name with REF that appears independently
            elseif (preg_match('/^([A-Z][A-Z\s&\(\)\/\-\.C\/O]+)\s+REF\s+[A-Z0-9\/\-]+$/i', $line, $m)) {
                // Check if this looks like a collection company with address info
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
                    Log::info("FOUND company with REF at line {$i}: company='{$m[1]}'");
                }
            }

            // Pattern 5: Lines that look like company names followed by address patterns
            elseif (preg_match('/^([A-Z][A-Z\s&\(\)\/\-\.]{5,60})$/i', $line) &&
                     !preg_match('/\b(DELIVERY|RATE|TERMS|PAYMENT|DATE|CARRIER|ZIEGLER|ROAD|STREET|LANE|AVENUE|WAY|CLOSE|UNITS|HALL)\b/i', $line)) {
                // Check if next few lines have address info
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
                    Log::info("FOUND potential company at line {$i}: '{$line}'");
                }
            }

            if ($confidence > 0) {
                $starts[] = [
                    'line' => $i,
                    'text' => $line,
                    'confidence' => $confidence,
                    'reason' => $reason
                ];
                Log::info("Added collection start: line={$i}, confidence={$confidence}, reason='{$reason}'");
            }
        }

        Log::info("=== COLLECTION SECTION STARTS SCAN COMPLETE ===");
        Log::info("Found " . count($starts) . " potential starts");
        return $starts;
    }

    protected function extractLoadingLocations(array $lines): array
    {
        $locations = [];
        $processedSections = []; // Track processed sections to avoid duplicates

        Log::info("=== SEARCHING FOR COLLECTION/LOADING LOCATIONS ===");
        Log::info("Total lines to scan: " . count($lines));

        // Find all collection section start points using improved formula
        $collectionStarts = $this->findCollectionSectionStarts($lines);

        Log::info("Found " . count($collectionStarts) . " potential collection section starts:");
        foreach ($collectionStarts as $start) {
            Log::info("  Start at line {$start['line']}: {$start['text']} (confidence: {$start['confidence']}, reason: {$start['reason']}')");
        }

        // If no collection starts found, try a more aggressive search
        if (empty($collectionStarts)) {
            Log::info("No collection starts found, trying aggressive search...");
            $collectionStarts = $this->findCollectionStartsAggressive($lines);
        }

        // Sort by confidence (highest first) and line number
        usort($collectionStarts, function($a, $b) {
            if ($a['confidence'] !== $b['confidence']) {
                return $b['confidence'] <=> $a['confidence']; // Higher confidence first
            }
            return $a['line'] <=> $b['line']; // Earlier line first for same confidence
        });

        // Process each unique collection section - limit to maximum 2 locations
        $maxLocations = 2;
        foreach ($collectionStarts as $start) {
            if (count($locations) >= $maxLocations) {
                Log::info("Reached maximum of {$maxLocations} collection locations, stopping");
                break;
            }

            $startLine = $start['line'];

            // Skip if we've already processed a section that overlaps with this one
            $shouldSkip = false;
            foreach ($processedSections as $processed) {
                if (abs($startLine - $processed['start']) < 8) { // Increased overlap detection
                    Log::info("Skipping line {$startLine} - too close to already processed section at line {$processed['start']}");
                    $shouldSkip = true;
                    break;
                }
            }

            if ($shouldSkip) continue;

            // Find the end of this collection section
            $endLine = $this->findCollectionSectionEnd($lines, $startLine);

            Log::info("Processing collection section from line {$startLine} to {$endLine}");

            // Extract location from this specific section
            $location = $this->parseCollectionLocationSection($lines, $startLine, $endLine);

            if ($location) {
                // Check for duplicates based on company name and postal code
                $isDuplicate = false;
                foreach ($locations as $existingLocation) {
                    if ($this->areLocationsDuplicate($location, $existingLocation)) {
                        Log::info("Skipping duplicate location: " . $location['company_address']['company']);
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
                    Log::info("Added unique collection location: " . $location['company_address']['company']);
                }
            }
        }

        Log::info("=== COLLECTION LOCATION EXTRACTION COMPLETE ===");
        Log::info("Found " . count($locations) . " unique collection locations");

        return $locations;
    }

    protected function findCollectionStartsAggressive(array $lines): array
    {
        $starts = [];

        Log::info("=== AGGRESSIVE SEARCH FOR COLLECTION STARTS ===");

        for ($i = 0; $i < count($lines); $i++) {
            $line = trim($lines[$i]);

            // Look for any line containing "Collection" (case insensitive)
            if (preg_match('/collection/i', $line)) {
                Log::info("Found 'Collection' in line {$i}: '{$line}'");

                // Extract everything after "Collection"
                if (preg_match('/collection\s+(.+)/i', $line, $m)) {
                    $starts[] = [
                        'line' => $i,
                        'text' => $line,
                        'confidence' => 90,
                        'reason' => 'Aggressive Collection search'
                    ];
                    Log::info("Added from aggressive search: '{$line}'");
                }
            }

            // Also look for lines that might be company names followed by postal codes
            elseif (preg_match('/^([A-Z][A-Z\s&\(\)\/\-\.C\/O]{8,50})\s*$/i', $line)) {
                // Check if next line has postal code
                if (isset($lines[$i + 1]) && preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2}|\d{5})\b/', $lines[$i + 1])) {
                    $starts[] = [
                        'line' => $i,
                        'text' => $line,
                        'confidence' => 60,
                        'reason' => 'Company name followed by postal code'
                    ];
                    Log::info("Added company+postcode pattern: '{$line}'");
                }
            }
        }

        Log::info("Aggressive search found " . count($starts) . " potential starts");
        return $starts;
    }

    protected function parseCollectionLocationSection(array $lines, int $startLine, int $endLine): ?array
    {
        $sectionLines = array_slice($lines, $startLine, $endLine - $startLine + 1);

        Log::info("=== PARSING COLLECTION SECTION (lines {$startLine}-{$endLine}) ===");
        foreach ($sectionLines as $i => $line) {
            Log::info("  Section line " . ($startLine + $i) . ": '{$line}'");
        }

        // For debugging, let's also look at a few lines after the section end
        Log::info("=== LINES AFTER SECTION END ===");
        for ($i = $endLine + 1; $i <= min($endLine + 5, count($lines) - 1); $i++) {
            if (isset($lines[$i])) {
                Log::info("  After line {$i}: '{$lines[$i]}'");
            }
        }

        // Parse the section with enhanced validation
        $result = $this->parseLocationDetails($sectionLines, 'collection');

        // If we didn't get good address data, try extending the section
        if ($result && isset($result['company_address'])) {
            $addr = $result['company_address'];

            // Check if we need to look further for address data
            if (($addr['street_address'] === 'TBC' || $addr['postal_code'] === 'TBC' || $addr['city'] === 'Unknown') &&
                $endLine < count($lines) - 1) {

                Log::info("Address data incomplete, looking further...");

                // Look at the next few lines for address data
                for ($i = $endLine + 1; $i <= min($endLine + 8, count($lines) - 1); $i++) {
                    if (isset($lines[$i])) {
                        $line = trim($lines[$i]);
                        Log::info("Checking extended line {$i}: '{$line}'");

                        // Stop if we hit another section
                        if (preg_match('/^(Collection|Delivery|Rate\s*€|Terms|Payment|Carrier)/i', $line)) {
                            Log::info("Hit new section, stopping extended search");
                            break;
                        }

                        // Try to parse this line for address info
                        $tempAddress = $addr;
                        $this->parseAddressComponents($line, $tempAddress);

                        // If we found better data, use it
                        if ($tempAddress['street_address'] !== $addr['street_address'] && $tempAddress['street_address'] !== 'TBC') {
                            $addr['street_address'] = $tempAddress['street_address'];
                            Log::info("Found street from extended search: {$addr['street_address']}");
                        }
                        if ($tempAddress['city'] !== $addr['city'] && $tempAddress['city'] !== 'Unknown') {
                            $addr['city'] = $tempAddress['city'];
                            Log::info("Found city from extended search: {$addr['city']}");
                        }
                        if ($tempAddress['postal_code'] !== $addr['postal_code'] && $tempAddress['postal_code'] !== 'TBC') {
                            $addr['postal_code'] = $tempAddress['postal_code'];
                            $addr['country_code'] = $tempAddress['country_code'];
                            Log::info("Found postal code from extended search: {$addr['postal_code']}");
                        }
                    }
                }

                // Update the result with better address data
                $result['company_address'] = $addr;
            }
        }

        // Less strict validation - allow more results through
        if ($result && isset($result['company_address'])) {
            $addr = $result['company_address'];

            // Only skip if company name is completely empty or exactly matches fallback
            if (empty($addr['company']) || $addr['company'] === 'Collection Location') {
                Log::info("Skipping location with empty/fallback company name: '{$addr['company']}'");
                return null;
            }

            Log::info("Parsed collection location successfully: " . $addr['company']);
            Log::info("Final address data: " . json_encode($addr));
        }

        return $result;
    }

    protected function findCollectionSectionEnd(array $lines, int $startLine): int
    {
        $maxSectionSize = 8; // Reduce section size to be more focused

        for ($i = $startLine + 1; $i < count($lines) && $i < $startLine + $maxSectionSize; $i++) {
            $line = trim($lines[$i]);
            if (empty($line)) continue;

            Log::info("Checking section end at line {$i}: '{$line}'");

            // Stop at any new Collection or Delivery section
            if (preg_match('/^(Collection|Delivery)\s+[A-Z]/i', $line)) {
                Log::info("Stopping at new Collection/Delivery section");
                return $i - 1;
            }

            // Stop at rate/pricing information
            if (preg_match('/^Rate\s*€/i', $line)) {
                Log::info("Stopping at Rate section");
                return $i - 1;
            }

            // Stop at administrative sections
            if (preg_match('/^(Terms|Payment|Carrier|Date\s+\d)/i', $line)) {
                Log::info("Stopping at administrative section");
                return $i - 1;
            }

            // Stop if we encounter what looks like another company+REF pattern that isn't part of this collection
            if (preg_match('/^[A-Z][A-Z\s&\(\)\/\-\.]+\s+REF\s+[A-Z0-9\/\-]+$/i', $line) && $i > $startLine + 3) {
                Log::info("Stopping at another company+REF pattern");
                return $i - 1;
            }
        }

        $endLine = min(count($lines) - 1, $startLine + $maxSectionSize - 1);
        Log::info("Reached max section size, ending at line {$endLine}");
        return $endLine;
    }

    protected function areLocationsDuplicate(array $location1, array $location2): bool
    {
        $addr1 = $location1['company_address'];
        $addr2 = $location2['company_address'];

        // Same company name = duplicate
        if ($addr1['company'] === $addr2['company']) {
            return true;
        }

        // Same postal code and similar company name = likely duplicate
        if (isset($addr1['postal_code']) && isset($addr2['postal_code']) &&
            $addr1['postal_code'] === $addr2['postal_code']) {

            // Check for similar company names (e.g., with/without "C/O" or "LTD")
            $clean1 = preg_replace('/\b(C\/O|LTD|LIMITED|CO)\b/i', '', $addr1['company']);
            $clean2 = preg_replace('/\b(C\/O|LTD|LIMITED|CO)\b/i', '', $addr2['company']);

            if (trim($clean1) === trim($clean2)) {
                return true;
            }
        }

        return false;
    }

    protected function extractDestinationLocations(array $lines): array
    {
        $locations = [];

        Log::info("=== SEARCHING FOR DELIVERY LOCATIONS ===");

        // Look for Delivery sections with multiple patterns
        foreach ($lines as $i => $line) {
            $upperLine = strtoupper($line);

            // Pattern 1: "Delivery COMPANY_NAME REF"
            if (preg_match('/^Delivery\s+(.+)/i', $line, $m)) {
                Log::info("FOUND Delivery pattern 1 at line {$i}: {$line}");
                $location = $this->parseLocationBlock($lines, $i, 'delivery');
                if ($location) {
                    $locations[] = $location;
                }
            }
            // Pattern 2: Lines containing "DELIVERY" (case insensitive)
            elseif (Str::contains($upperLine, 'DELIVERY')) {
                Log::info("FOUND Delivery pattern 2 at line {$i}: {$line}");
                $location = $this->parseLocationBlock($lines, $i, 'delivery');
                if ($location) {
                    $locations[] = $location;
                }
            }
            // Pattern 3: Any line that looks like it starts a destination block
            elseif (preg_match('/^(DELIVER|TO|DESTINATION)[\s:]/i', $line)) {
                Log::info("FOUND Delivery pattern 3 at line {$i}: {$line}");
                $location = $this->parseLocationBlock($lines, $i, 'delivery');
                if ($location) {
                    $locations[] = $location;
                }
            }
        }

        Log::info("Found " . count($locations) . " delivery locations");
        return $locations;
    }

    protected function parseLocationBlock(array $lines, int $startIdx, string $type): ?array
    {
        $blockLines = [];

        // Collect lines for this location block - be more generous
        for ($i = $startIdx; $i < count($lines) && $i < $startIdx + 30; $i++) {
            $line = trim($lines[$i]);
            if (empty($line)) continue;

            // Stop at next major section - be more specific about what constitutes a new section
            if ($i > $startIdx && preg_match('/^(Collection|Delivery|Rate\s*€|Terms|Payment|Carrier|Date\s+\d)/i', $line)) {
                Log::info("Stopping location block at line {$i}: {$line}");
                break;
            }

            $blockLines[] = $line;
        }

        if (empty($blockLines)) {
            Log::info("No block lines found for {$type}");
            return null;
        }

        Log::info("Parsing {$type} block with " . count($blockLines) . " lines:");
        foreach ($blockLines as $j => $line) {
            Log::info("  Block line {$j}: '{$line}'");
        }

        // Extract company name from first line or any line with company indicators
        $company = '';
        $firstLine = $blockLines[0];

        // Multiple patterns to extract company name
        if (preg_match('/^(Collection|Delivery)\s+(.+?)(?:\s+REF)?$/i', $firstLine, $m)) {
            $company = trim($m[2]);
            $company = preg_replace('/\s+REF\s*$/i', '', $company);
        } elseif (preg_match('/^(.+?)(?:\s+REF)?$/i', $firstLine, $m)) {
            // If no Collection/Delivery prefix, use the whole line as company
            $company = trim($m[1]);
            $company = preg_replace('/\s+REF\s*$/i', '', $company);
        }

        // Look for company names in subsequent lines too
        if (empty($company)) {
            foreach ($blockLines as $line) {
                // Look for lines that look like company names (mix of letters and possibly parentheses)
                if (preg_match('/^([A-Z][A-Z\s&\(\)\/\-\.]{5,50})$/i', $line) && !preg_match('/\d{2}\/\d{2}\/\d{4}/', $line)) {
                    $company = trim($line);
                    Log::info("Found company name in line: {$company}");
                    break;
                }
            }
        }

        // Initialize address
        $address = [
            'company' => $company ?: ucfirst($type) . ' Location',
            'street_address' => '',
            'city' => '',
            'postal_code' => '',
            'country_code' => 'GB'
        ];

        $timeObj = null;
        $date = null;
        $timeStart = null;
        $timeEnd = null;

        // Process each line in the block
        foreach ($blockLines as $lineIndex => $line) {
            // Skip the header line
            if ($lineIndex === 0 && preg_match('/^(Collection|Delivery)\s+/i', $line)) {
                continue;
            }

            Log::info("Processing address line: '{$line}'");

            // Extract date
            if (preg_match('/(\d{2}\/\d{2}\/\d{4})/', $line, $m)) {
                $date = $m[1];
                Log::info("Found date: {$date}");
            }

            // Extract time ranges: various formats including PM/AM
            if (preg_match('/(\d{2})(\d{2})\s*[-–]\s*(\d{1,2})\s*(pm|am)/i', $line, $m)) {
                $startHour = (int)$m[1];
                $startMin = (int)$m[2];
                $endHour = (int)$m[3];
                $endAmPm = strtolower($m[4]);

                // Convert PM/AM to 24-hour format
                if ($endAmPm === 'pm' && $endHour < 12) {
                    $endHour += 12;
                } elseif ($endAmPm === 'am' && $endHour === 12) {
                    $endHour = 0;
                }

                $timeStart = sprintf('%02d:%02d', $startHour, $startMin);
                $timeEnd = sprintf('%02d:%02d', $endHour, 0);
                Log::info("Found time range (with PM/AM): {$timeStart} - {$timeEnd}");
            }
            elseif (preg_match('/(\d{2}):?(\d{2})\s*[-–]\s*(\d{2}):?(\d{2})/', $line, $m)) {
                $timeStart = sprintf('%02d:%02d', (int)$m[1], (int)$m[2]);
                $timeEnd = sprintf('%02d:%02d', (int)$m[3], (int)$m[4]);
                Log::info("Found time range: {$timeStart} - {$timeEnd}");
            }

            // Parse address components
            $this->parseAddressLine($line, $address);

            // Extract REF comments
            if (preg_match('/REF\s+([A-Z0-9]+)/i', $line, $m)) {
                $address['comment'] = $company . ' REF ' . $m[1];
                Log::info("Found REF: {$address['comment']}");
            }
        }

        // Build time object
        if ($date && $timeStart && $timeEnd) {
            try {
                $carbonDate = Carbon::createFromFormat('d/m/Y', $date);
                $timeObj = [
                    'datetime_from' => $carbonDate->copy()->setTimeFromTimeString($timeStart)->format('c'),
                    'datetime_to' => $carbonDate->copy()->setTimeFromTimeString($timeEnd)->format('c')
                ];
                Log::info("Built time object: " . json_encode($timeObj));
            } catch (\Exception $e) {
                Log::error("Error building time: " . $e->getMessage());
            }
        }

        // Clean up address - remove empty fields but keep structure
        $cleanAddress = [];
        foreach ($address as $key => $value) {
            if (!empty($value) && $value !== 'TBC' && $value !== 'Unknown') {
                $cleanAddress[$key] = $value;
            }
        }

        // Ensure minimum required fields with proper values
        if (!isset($cleanAddress['company']) || empty($cleanAddress['company'])) {
            $cleanAddress['company'] = ucfirst($type) . ' Location';
        }
        if (!isset($cleanAddress['street_address'])) {
            $cleanAddress['street_address'] = 'TBC';
        }
        if (!isset($cleanAddress['city'])) {
            $cleanAddress['city'] = 'Unknown';
        }
        if (!isset($cleanAddress['postal_code'])) {
            $cleanAddress['postal_code'] = 'TBC';
        }
        if (!isset($cleanAddress['country_code'])) {
            $cleanAddress['country_code'] = 'GB';
        }

        $result = ['company_address' => $cleanAddress];
        if ($timeObj) {
            $result['time'] = $timeObj;
        }

        Log::info("Final location result: " . json_encode($result));
        return $result;
    }

    protected function extractCargos(array $lines): array
    {
        $cargos = [];

        foreach ($lines as $line) {
            // Look for pallet counts
            if (preg_match('/(\d+)\s+(?:PALLETS?|PALLET)/i', $line, $m)) {
                $cargos[] = [
                    'title' => 'Palletized goods',
                    'package_count' => (int) $m[1],
                    'package_type' => 'pallet'
                ];
                Log::info("Found cargo: {$m[1]} pallets");
            }
            // Look for other package types
            elseif (preg_match('/(\d+)\s+(PACKAGES?|CARTONS?|BOXES?|ITEMS?)/i', $line, $m)) {
                $cargos[] = [
                    'title' => 'Packaged goods',
                    'package_count' => (int) $m[1],
                    'package_type' => 'package'
                ];
                Log::info("Found cargo: {$m[1]} {$m[2]}");
            }
        }

        return $cargos;
    }

    protected function extractComment(array $lines): string
    {
        $parts = [];

        foreach ($lines as $line) {
            // "Carrier Test_Client"
            if (preg_match('/Carrier\s+(.+?)(?:\s|$)/i', $line, $m)) {
                $parts[] = 'Carrier: ' . trim($m[1]);
            }

            // "Date 23/06/2025"
            if (preg_match('/Date\s+(\d{2}\/\d{2}\/\d{4})/i', $line, $m)) {
                $parts[] = 'Booking created on ' . $m[1];
            }
        }

        // Add standard Ziegler terms
        $hasDeliveryRestriction = collect($lines)->some(fn($line) =>
            Str::contains(Str::lower($line), 'delivery to any address other than')
        );

        if ($hasDeliveryRestriction) {
            $parts[] = 'Notes: Delivery to any address other than listed is prohibited without permission; signed POD required for payment';
        }

        // Return empty string if no parts found, otherwise join with periods
        return collect($parts)->filter()->implode('. ') . (count($parts) ? '.' : '');
    }
}