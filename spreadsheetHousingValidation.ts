function main(workbook: ExcelScript.Workbook): void {
    const sheet = workbook.getWorksheet("Sheet1"); //Make sure copy the data to Sheet1 or rename the sheet name
    const usedRange = sheet.getUsedRange();

    // Use getTexts() to get all cell values as strings
    const rows: string[][] = usedRange.getTexts();

    // Define the type for field settings
    type FieldSetting = {
        type: 'text' | 'numeric' | 'date' | 'bigint';
        maxLength?: number;
        nullable?: boolean;
        special?: string;
        allowedValues?: string[];
    };

    // Update settings to include all fields
    const settings: { [key: string]: FieldSetting } = {
        'MINISTRY_CODE': { type: 'text', maxLength: 50 },
        'BUSINESS_AREA': { type: 'text', maxLength: 100 },
        'SOURCE_SYSTEM_ACRONYM': { type: 'text', maxLength: 10, nullable: true },
        'PERMIT_TYPE': { type: 'text', maxLength: 100 },
        'PROJECT_ID': { type: 'text', maxLength: 20 },
        'APPLICATION_ID': { type: 'text', maxLength: 20 },
        'PROJECT_NAME': { type: 'text', maxLength: 255, nullable: true },
        'PROJECT_DESCRIPTION': { type: 'text', nullable: true },
        'PROJECT_LOCATION': { type: 'text', maxLength: 255 },
        'UTM_EASTING': { type: 'numeric', nullable: true, special: 'na' },
        'UTM_NORTHING': { type: 'numeric', nullable: true, special: 'na' },
        'RECEIVED_DATE': { type: 'date', nullable: false },
        'ACCEPTED_DATE': { type: 'date', nullable: true },
        'ADJUDICATION_DATE': { type: 'date', nullable: true },
        'REJECTED_DATE': { type: 'date', nullable: true },
        'AMENDMENT_RENEWAL_DATE': { type: 'date', nullable: true },
        'TECH_REVIEW_COMPLETION_DATE': { type: 'date', nullable: true },
        'FN_CONSULTN_START_DATE': { type: 'date', nullable: true },
        'FN_CONSULTN_COMPLETION_DATE': { type: 'date', nullable: true },
        'FN_CONSULTN_COMMENT': { type: 'text', nullable: true },
        'REGION_NAME': { type: 'text', maxLength: 100 },
        'INDIGENOUS_LED_IND': {
            type: 'text',
            maxLength: 7,
            nullable: true,
            allowedValues: ['Y', 'N', 'Unknown'],
        },
        'RENTAL_LICENSE_IND': {
            type: 'text',
            maxLength: 7,
            nullable: true,
            allowedValues: ['Y', 'N', 'Unknown'],
        },
        'SOCIAL_HOUSING_IND': {
            type: 'text',
            maxLength: 7,
            nullable: true,
            allowedValues: ['Y', 'N', 'Unknown'],
        },
        'HOUSING_TYPE': { type: 'text', maxLength: 100 },
        'ESTIMATED_HOUSING': { type: 'numeric', nullable: true },
        'APPLICATION_STATUS': { type: 'text', maxLength: 50 },
        'BUSINESS_AREA_FILE_NUMBER': { type: 'bigint', nullable: true },
    };

    let allErrors: string[] = [];

    // Function to normalize strings by removing non-printable characters and trimming
    function normalizeString(value: string): string {
        return value
            .replace(/[\u200B-\u200D\uFEFF]/g, '') // Remove zero-width spaces and similar characters
            .trim()
            .replace(/[ \t]+/g, ' '); // Replace multiple spaces or tabs with a single space
    }

    // Get the header row and build a mapping from field names to column indices
    const headerRow = rows[0];

    const fieldIndexMap: { [key: string]: number } = {};
    headerRow.forEach((headerName: string, colIndex: number) => {
        const normalizedHeaderName = headerName.replace(/\s+/g, ''); // Remove all whitespace
        fieldIndexMap[normalizedHeaderName] = colIndex;
    });

    // Normalize the keys in the settings to match the normalized header names
    const normalizedSettings: { [key: string]: FieldSetting } = {};
    Object.keys(settings).forEach((key) => {
        const normalizedKey = key.replace(/\s+/g, '');
        normalizedSettings[normalizedKey] = settings[key];
    });

    // Validate each row
    for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
        // Skip header row
        const row = rows[rowIndex];
        let errors: string[] = [];
        let receivedDate: Date | null = null;

        // Get RECEIVED_DATE value
        const receivedDateIndex = fieldIndexMap['RECEIVED_DATE'];
        let receivedDateValue = row[receivedDateIndex]?.trim();

        if (receivedDateValue !== undefined && receivedDateValue !== '') {
            // Validate date format
            if (!isValidDateFormat(receivedDateValue)) {
                errors.push(
                    `Error in RECEIVED_DATE at Row ${rowIndex + 1
                    }: Date must be in 'YYYY-MM-DD' format.`
                );
            } else {
                receivedDate = parseDate(receivedDateValue);
                if (!receivedDate) {
                    errors.push(
                        `Error in RECEIVED_DATE at Row ${rowIndex + 1}: Must be a valid date.`
                    );
                }
            }
        } else {
            errors.push(
                `Error in RECEIVED_DATE at Row ${rowIndex + 1}: Field cannot be empty.`
            );
        }

        Object.keys(normalizedSettings).forEach((normalizedKey: string) => {
            if (normalizedKey === 'RECEIVED_DATE') return; // Already validated above

            const setting: FieldSetting = normalizedSettings[normalizedKey];
            const colIndex = fieldIndexMap[normalizedKey];

            if (colIndex === undefined) {
                // Column not found
                errors.push(`Error: Column '${normalizedKey}' not found in data.`);
                return;
            }

            let cellValue = row[colIndex]?.trim();

            // Apply normalization to cell values
            if (cellValue) {
                cellValue = normalizeString(cellValue);
            }

            // Check for empty values in non-nullable fields
            if (!setting.nullable && (cellValue === undefined || cellValue === '')) {
                errors.push(
                    `Error in ${normalizedKey} at Row ${rowIndex + 1
                    }: Field cannot be empty.`
                );
                return; // Skip further validation for this field
            }

            // If field is nullable and value is empty, skip validation
            if (setting.nullable && (cellValue === undefined || cellValue === '')) {
                return;
            }

            // Check data types and constraints
            switch (setting.type) {
                case 'text':
                    const textValue = cellValue;

                    if (!setting.nullable && textValue === '') {
                        errors.push(
                            `Error in ${normalizedKey} at Row ${rowIndex + 1
                            }: Field cannot be empty.`
                        );
                        return;
                    }

                    if (setting.maxLength && textValue.length > setting.maxLength) {
                        errors.push(
                            `Error in ${normalizedKey} at Row ${rowIndex + 1
                            }: Text exceeds maximum length of ${setting.maxLength}.`
                        );
                    }

                    if (
                        setting.allowedValues &&
                        !setting.allowedValues.includes(textValue)
                    ) {
                        errors.push(
                            `Error in ${normalizedKey} at Row ${rowIndex + 1
                            }: Value must be one of [${setting.allowedValues.join(', ')}].`
                        );
                    }
                    break;
                case 'numeric':
                    if (
                        setting.special &&
                        cellValue.toLowerCase() === setting.special.toLowerCase()
                    ) {
                        // Accept special value
                    } else if (!isNaN(Number(cellValue))) {
                        // Valid number in string form
                    } else {
                        errors.push(
                            `Error in ${normalizedKey} at Row ${rowIndex + 1
                            }: Must be numeric or special value '${setting.special}'.`
                        );
                    }
                    break;
                case 'date':
                    // Validate date format
                    if (!isValidDateFormat(cellValue)) {
                        errors.push(
                            `Error in ${normalizedKey} at Row ${rowIndex + 1
                            }: Date must be in 'YYYY-MM-DD' format.`
                        );
                    } else {
                        const dateValue = parseDate(cellValue);
                        if (!dateValue) {
                            errors.push(
                                `Error in ${normalizedKey} at Row ${rowIndex + 1
                                }: Must be a valid date.`
                            );
                        } else if (receivedDate && dateValue < receivedDate) {
                            errors.push(
                                `Error in ${normalizedKey} at Row ${rowIndex + 1
                                }: Date is earlier than RECEIVED_DATE.`
                            );
                        }
                    }
                    break;
                case 'bigint':
                    if (/^-?\d+$/.test(cellValue)) {
                        // Valid integer
                    } else {
                        errors.push(
                            `Error in ${normalizedKey} at Row ${rowIndex + 1
                            }: Must be a valid integer.`
                        );
                    }
                    break;
            }
        });

        // Log all errors associated with an APPLICATION_ID
        if (errors.length > 0) {
            const applicationIdIndex = fieldIndexMap['APPLICATION_ID'];
            const applicationId = row[applicationIdIndex]?.trim();
            allErrors.push(
                `Errors for APPLICATION_ID ${applicationId} at Row ${rowIndex + 1}:`
            );
            allErrors = allErrors.concat(errors);
        }
    }

    // Display all errors or success message
    if (allErrors.length > 0) {
        console.log(allErrors.join('\n'));
    } else {
        console.log('Great! Data validated. Ready to use.');
    }

    // Helper function to parse date values
    function parseDate(value: string): Date | null {
        const date = new Date(value);
        if (!isNaN(date.getTime())) {
            return date;
        }
        return null;
    }

    // Helper function to validate date format 'YYYY-MM-DD'
    function isValidDateFormat(value: string): boolean {
        const dateFormatRegex = /^\d{4}-\d{2}-\d{2}$/;
        return dateFormatRegex.test(value);
    }
}
