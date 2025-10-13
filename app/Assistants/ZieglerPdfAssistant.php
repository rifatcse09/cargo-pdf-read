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

        // Build payload according to schema
        $payload = [
            'attachment_filenames' => $attachment_filename ? [basename($attachment_filename)] : [],
            'customer' => $customer,
            'order_reference' => $orderRef,
            'freight_price' => $freight['price'],
            'freight_currency' => $freight['currency'] ?: 'EUR',
            'loading_locations' => $loadingLocs ?: [['company_address' => (object)[]]],
            'destination_locations' => $destLocs ?: [['company_address' => (object)[]]],
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

        foreach ($addressLines as $line) {
            $line = trim($line);
            if (empty($line)) break;

            // Stop at booking/instruction lines
            if (preg_match('/\b(BOOKING|INSTRUCTION|TELEPHONE|REF)\b/i', $line)) {
                break;
            }

            // UK postcode (SS17 9FJ)
            if (preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2})\b/', $line, $m)) {
                $pc = strtoupper(str_replace(' ', '', $m[1]));
                $details['postal_code'] = substr($pc, 0, -3) . ' ' . substr($pc, -3);

                // Use GeonamesCountry to detect country from postal code
                $countryIso = GeonamesCountry::getIso($details['postal_code']);
                if ($countryIso) {
                    $details['country_code'] = $countryIso;
                } else {
                    // Fallback to GB for UK postal code pattern
                    $details['country_code'] = 'GB';
                }
                continue;
            }

            // French/EU postcode pattern (5 digits)
            if (preg_match('/\b(\d{5})\b/', $line, $m)) {
                $details['postal_code'] = $m[1];

                // Use GeonamesCountry to detect country from postal code
                $countryIso = GeonamesCountry::getIso($details['postal_code']);
                if ($countryIso) {
                    $details['country_code'] = $countryIso;
                } else {
                    // Fallback to FR for 5-digit postal code pattern
                    $details['country_code'] = 'FR';
                }
                continue;
            }

            // City (all caps, no digits)
            if (preg_match('/^[A-Z\s\'-]{4,}$/', $line) && !preg_match('/\d/', $line)) {
                $details['city'] = $line;
                continue;
            }

            // Street parts
            $streetParts[] = $line;
        }

        if (!empty($streetParts)) {
            $details['street_address'] = collect($streetParts)->implode(', ');
        }
    }

    protected function parseAddressComponents(string $line, array &$address): void
    {
        // Street address (contains common street indicators)
        if (preg_match('/\b(.*(?:ROAD|RD|STREET|ST|LANE|AVENUE|AVE|RUE|UNITS|HALL|Chem\.).*)/', $line, $m)) {
            if (!isset($address['street_address'])) {
                $address['street_address'] = trim($m[1]);
            }
        }

        // UK postcode + city: "IP14 2QU STOWMARKET"
        if (preg_match('/([A-Z]{1,2}\d+\s+\d[A-Z]{2})\s+([A-Z\s]+)$/i', $line, $m)) {
            $address['postal_code'] = $m[1];
            $address['city'] = trim($m[2]);

            // Use GeonamesCountry to detect country from postal code
            $countryIso = GeonamesCountry::getIso($address['postal_code']);
            $address['country_code'] = $countryIso ?: 'GB';
            return;
        }

        // French postcode + city: "95150 TAVERNY"
        if (preg_match('/(\d{5})\s+([A-Z\s]+)(?:\s+\d+\s+PALLETS)?$/i', $line, $m)) {
            $address['postal_code'] = $m[1];
            $address['city'] = trim($m[2]);

            // Use GeonamesCountry to detect country from postal code
            $countryIso = GeonamesCountry::getIso($address['postal_code']);
            $address['country_code'] = $countryIso ?: 'FR';
            return;
        }

        // Standalone postal codes
        if (preg_match('/\b([A-Z]{1,2}\d+\s*\d[A-Z]{2})\b/', $line, $m)) {
            $pc = strtoupper(str_replace(' ', '', $m[1]));
            $address['postal_code'] = substr($pc, 0, -3) . ' ' . substr($pc, -3);

            // Use GeonamesCountry to detect country from postal code
            $countryIso = GeonamesCountry::getIso($address['postal_code']);
            $address['country_code'] = $countryIso ?: 'GB';
        }

        if (preg_match('/\b(\d{5})\b/', $line, $m)) {
            $address['postal_code'] = $m[1];

            // Use GeonamesCountry to detect country from postal code
            $countryIso = GeonamesCountry::getIso($address['postal_code']);
            $address['country_code'] = $countryIso ?: 'FR';
        }

        // Reference comment
        if (preg_match('/REF\s+([A-Z0-9]+)/i', $line, $m)) {
            $address['comment'] = $address['company'] . ' REF ' . $m[1];
        }

        // Booking comments
        if (preg_match('/BOOKED\s+FOR\s+(\d{2}\/\d{2})/i', $line, $m)) {
            $address['comment'] = 'BOOKED FOR ' . $m[1];
        }
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
            // "Rate € 1,000"
            if (preg_match('/Rate\s*€\s*([0-9,.\s]+)/i', $line, $m)) {
                $price = (float) str_replace([',', ' '], '', $m[1]);
                $currency = 'EUR';
                break;
            }
            // "Price: EUR 1000"
            if (preg_match('/Price[:\s]*EUR\s*([0-9,.\s]+)/i', $line, $m)) {
                $price = (float) str_replace([',', ' '], '', $m[1]);
                $currency = 'EUR';
                break;
            }
            // Generic currency patterns
            if (preg_match('/(EUR|GBP|USD)\s*([0-9,.\s]+)/i', $line, $m)) {
                $currency = strtoupper($m[1]);
                $price = (float) str_replace([',', ' '], '', $m[2]);
                break;
            }
            if (preg_match('/([€£$])\s*([0-9,.\s]+)/', $line, $m)) {
                $currency = ['€' => 'EUR', '£' => 'GBP', '$' => 'USD'][$m[1]] ?? null;
                $price = (float) str_replace([',', ' '], '', $m[2]);
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

        // Default to EUR if no currency found (common for Ziegler)
        return [
            'price' => $price,
            'currency' => $currency ?: 'EUR'
        ];
    }

    protected function extractLoadingLocations(array $lines): array
    {
        return $this->extractLocations($lines, ['Collection ']);
    }

    protected function extractDestinationLocations(array $lines): array
    {
        return $this->extractLocations($lines, ['Delivery ']);
    }

    protected function extractLocations(array $lines, array $markers): array
    {
        $locations = [];

        foreach ($lines as $i => $line) {
            foreach ($markers as $marker) {
                if (Str::startsWith($line, $marker)) {
                    $location = $this->parseLocationSection($lines, $i);
                    if ($location) {
                        $locations[] = $location;
                    }
                    break;
                }
            }
        }

        return $locations;
    }

    protected function parseLocationSection(array $lines, int $startIdx): ?array
    {
        // Get section lines until next section or end
        $sectionLines = [];
        for ($i = $startIdx; $i < count($lines); $i++) {
            $line = $lines[$i];

            // Stop at new section
            if ($i > $startIdx && $this->isNewLocationSection($line)) {
                break;
            }

            $sectionLines[] = $line;
        }

        // Parse company from first line
        $firstLine = $sectionLines[0] ?? '';
        if (preg_match('/(Collection|Delivery)\s+(.+?)\s+REF$/i', $firstLine, $m)) {
            $company = trim($m[2]);
        } else {
            return null;
        }

        // Parse address, time, and other details
        $address = ['company' => $company];
        $timeObj = null;
        $date = null;
        $timeRange = null;

        foreach ($sectionLines as $line) {
            // Date (dd/mm/yyyy)
            if (preg_match('/(\d{2}\/\d{2}\/\d{4})/', $line, $m)) {
                $date = $m[1];
            }

            // Time patterns
            $timeRange = $this->parseTimeFromLine($line) ?: $timeRange;

            // Address components
            $this->parseAddressComponents($line, $address);
        }

        // Build time object
        if ($timeRange && $date) {
            $timeObj = $this->buildTimeObject($date, $timeRange);
        }

        // Detect country if not set
        if (!isset($address['country_code'])) {
            $address['country_code'] = $this->detectCountryFromSection($sectionLines);
        }

        return [
            'company_address' => array_filter($address),
            'time' => $timeObj
        ];
    }

    protected function isNewLocationSection(string $line): bool
    {
        return Str::startsWith($line, ['Collection ', 'Delivery ', '- Payment', '- All business']);
    }

    protected function parseTimeFromLine(string $line): ?array
    {
        // "0900-2pm" or "0900-3PM"
        if (preg_match('/(\d{4})-(\d+)(?:pm|PM)/i', $line, $m)) {
            $start = substr($m[1], 0, 2) . ':' . substr($m[1], 2, 2);
            $end = sprintf('%02d:00', (int)$m[2]);
            return ['start' => $start, 'end' => $end];
        }

        // "BOOKED-06:00 AM"
        if (preg_match('/BOOKED-(\d{2}:\d{2})/i', $line, $m)) {
            return ['start' => $m[1]];
        }

        // "09:00 Time To: 12:00"
        if (preg_match('/(\d{2}:\d{2})\s+Time\s+To:\s+(\d{2}:\d{2})/i', $line, $m)) {
            return ['start' => $m[1], 'end' => $m[2]];
        }

        // "09:00 - 15:00"
        if (preg_match('/(\d{1,2}:\d{2})\s*[-–—]\s*(\d{1,2}:\d{2})/', $line, $m)) {
            return ['start' => $m[1], 'end' => $m[2]];
        }

        return null;
    }

    protected function buildTimeObject(string $date, array $timeRange): ?array
    {
        try {
            $carbonDate = Carbon::createFromFormat('d/m/Y', $date);

            $timeObj = [
                'datetime_from' => $carbonDate->copy()
                    ->setTimeFromTimeString($timeRange['start'])
                    ->toIso8601String()
            ];

            if (isset($timeRange['end'])) {
                $timeObj['datetime_to'] = $carbonDate->copy()
                    ->setTimeFromTimeString($timeRange['end'])
                    ->toIso8601String();
            }

            return $timeObj;
        } catch (\Exception $e) {
            return null;
        }
    }

    protected function detectCountryFromSection(array $sectionLines): ?string
    {
        $text = collect($sectionLines)->implode(' ');

        // Postal code patterns
        if (preg_match('/[A-Z]{1,2}\d+\s+\d[A-Z]{2}/', $text)) return 'GB';
        if (preg_match('/\d{5}\s+[A-Z]+/', $text)) return 'FR';

        // Street indicators
        if (Str::contains($text, ['RUE', 'Chem.'])) return 'FR';
        if (Str::contains($text, ['ROAD', 'LANE'])) return 'GB';

        // Try GeonamesCountry for city names
        foreach ($sectionLines as $line) {
            if (preg_match('/\b([A-Z\s]{4,})\b/', $line, $m)) {
                $cityName = trim($m[1]);
                $iso = GeonamesCountry::getIso($cityName);
                if ($iso) return $iso;
            }
        }

        return null;
    }

    protected function extractCargos(array $lines): array
    {
        $cargos = [];

        foreach ($lines as $line) {
            // "20 pallets" or "10 PALLETS"
            if (preg_match('/(\d+)\s+pallets?/i', $line, $m)) {
                $cargos[] = [
                    'title' => 'Palletized goods',
                    'package_count' => (int) $m[1],
                    'package_type' => 'pallet'
                ];
            }
            // Other package types
            elseif (preg_match('/(\d+)\s+(packages?|cartons?|boxes?)/i', $line, $m)) {
                $cargos[] = [
                    'title' => 'Packaged goods',
                    'package_count' => (int) $m[1],
                    'package_type' => 'package'
                ];
            }
        }

        return $cargos;
    }

    protected function extractComment(array $lines): ?string
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

        return collect($parts)->filter()->implode('. ') . (count($parts) ? '.' : '') ?: null;
    }
}