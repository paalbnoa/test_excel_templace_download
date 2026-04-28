export const NAME_ALLOWED_CHARACTERS =
  "ABCDEFGHIJKLMNOPQRSTUVWXYZ\u00C6\u00D8\u00C5abcdefghijklmnopqrstuvwxyz\u00E6\u00F8\u00E5";

const NORWEGIAN_PERSONAL_NUMBER_START_YEAR = 1854;
const NORWEGIAN_PERSONAL_NUMBER_END_YEAR = 2039;
const K1_WEIGHTS = [3, 7, 6, 1, 8, 9, 4, 5, 2];
const K2_WEIGHTS = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2];

function buildLettersOnlyFormula(cellReference) {
  const strippedValueExpression = [...NAME_ALLOWED_CHARACTERS].reduce(
    (expression, character) => `SUBSTITUTE(${expression},"${character}","")`,
    `TRIM(${cellReference})`
  );

  return `AND(LEN(TRIM(${cellReference}))>0,LEN(${strippedValueExpression})=0)`;
}

function buildPrefixFormula(cellReference) {
  const digitTailExpression = Array.from({ length: 10 }, (_, digit) => String(digit)).reduce(
    (expression, digit) => `SUBSTITUTE(${expression},"${digit}","")`,
    `RIGHT(${cellReference},LEN(${cellReference})-1)`
  );

  return `AND(LEFT(${cellReference},1)="+",LEN(${cellReference})>1,LEFT(${cellReference},3)<>"+00",COUNTIF(${cellReference},"* *")=0,LEN(${digitTailExpression})=0)`;
}

function normalizeText(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}

function normalizeNumberLikeText(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return typeof value === "number" ? String(value) : String(value).trim();
}

function calculateNorwegianCheckDigit(digits, weights) {
  const sum = weights.reduce((total, weight, index) => {
    return total + Number(digits[index]) * weight;
  }, 0);
  const checkDigit = 11 - (sum % 11);

  if (checkDigit === 11) {
    return 0;
  }

  if (checkDigit === 10) {
    return null;
  }

  return checkDigit;
}

function isRealDate(year, month, day) {
  const date = new Date(Date.UTC(year, month - 1, day));

  return (
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
  );
}

function isValidNorwegianIndividualNumberForYear(year, individualNumber) {
  if (
    year >= 1854 &&
    year <= 1899 &&
    individualNumber >= 500 &&
    individualNumber <= 749
  ) {
    return true;
  }

  if (year >= 1900 && year <= 1999 && individualNumber >= 0 && individualNumber <= 499) {
    return true;
  }

  if (year >= 1940 && year <= 1999 && individualNumber >= 900 && individualNumber <= 999) {
    return true;
  }

  return year >= 2000 && year <= 2039 && individualNumber >= 500 && individualNumber <= 999;
}

export function validateNorwegianPersonalNumber(value) {
  const normalized = normalizeNumberLikeText(value);

  if (!/^\d{11}$/.test(normalized)) {
    return false;
  }

  const day = Number(normalized.slice(0, 2));
  const month = Number(normalized.slice(2, 4));
  const shortYear = Number(normalized.slice(4, 6));
  const individualNumber = Number(normalized.slice(6, 9));
  const candidateYears = [1800 + shortYear, 1900 + shortYear, 2000 + shortYear];
  const hasValidBirthDateAndIndividualNumber = candidateYears.some((year) => {
    return (
      year >= NORWEGIAN_PERSONAL_NUMBER_START_YEAR &&
      year <= NORWEGIAN_PERSONAL_NUMBER_END_YEAR &&
      isRealDate(year, month, day) &&
      isValidNorwegianIndividualNumberForYear(year, individualNumber)
    );
  });

  if (!hasValidBirthDateAndIndividualNumber) {
    return false;
  }

  const k1 = calculateNorwegianCheckDigit(normalized.slice(0, 9), K1_WEIGHTS);
  if (k1 === null || k1 !== Number(normalized[9])) {
    return false;
  }

  const k2 = calculateNorwegianCheckDigit(normalized.slice(0, 10), K2_WEIGHTS);
  return k2 !== null && k2 === Number(normalized[10]);
}

export const FIELD_RULES = {
  PersonID: {
    instruction:
      "PersonID: This field is required and must be a valid 11-digit Norwegian personal number. The first 6 digits must be a real birth date from 01.01.1854 to 31.12.2039, followed by a valid individual number for the birth year and two valid MOD11 check digits.",
    errorTitle: "Invalid PersonID",
    errorMessage:
      "PersonID is required and must be a valid 11-digit Norwegian personal number with a real birth date, valid individual number, and valid MOD11 check digits.",
    excel: {
      type: "custom",
      allowBlank: false,
      formula: (cellReference) =>
        `AND(LEN(${cellReference}&"")=11,SUMPRODUCT(--ISNUMBER(FIND(MID(${cellReference}&"",ROW($1:$11),1),"0123456789")))=11)`
    },
    validate: (value) => validateNorwegianPersonalNumber(value)
  },
  Fornavn: {
    instruction: "Fornavn: This field is required and must contain letters only.",
    errorTitle: "Invalid Fornavn",
    errorMessage: "Fornavn is required and must contain letters only.",
    excel: {
      type: "custom",
      allowBlank: false,
      formula: (cellReference) => buildLettersOnlyFormula(cellReference)
    },
    validate: (value) => /^[A-Za-z\u00C6\u00D8\u00C5\u00E6\u00F8\u00E5]+$/u.test(normalizeText(value))
  },
  Etternavn: {
    instruction:
      "Etternavn: This field is required and must contain letters only. No special characters are allowed.",
    errorTitle: "Invalid Etternavn",
    errorMessage: "Etternavn is required and must contain letters only. No special characters are allowed.",
    excel: {
      type: "custom",
      allowBlank: false,
      formula: (cellReference) => buildLettersOnlyFormula(cellReference)
    },
    validate: (value) => /^[A-Za-z\u00C6\u00D8\u00C5\u00E6\u00F8\u00E5]+$/u.test(normalizeText(value))
  },
  "Fritatt.sem.avg": {
    instruction: 'Fritatt.sem.avg: This field may be blank, or contain only "Ja" or "Nei".',
    errorTitle: "Invalid Fritatt.sem.avg",
    errorMessage: 'Fritatt.sem.avg may only be blank, "Ja", or "Nei".',
    excel: {
      type: "custom",
      allowBlank: true,
      formula: (cellReference) =>
        `OR(LEN(TRIM(${cellReference}))=0,TRIM(${cellReference})="Ja",TRIM(${cellReference})="Nei")`
    },
    validate: (value) => {
      const normalized = normalizeText(value);
      return normalized === "" || normalized === "Ja" || normalized === "Nei";
    }
  },
  Epost: {
    instruction:
      "Epost: This field is required and must look like a valid email address and contain one @ sign and a period after the @ sign.",
    errorTitle: "Invalid email",
    errorMessage: "Epost is required and must be a valid email address.",
    excel: {
      type: "custom",
      allowBlank: false,
      formula: (cellReference) =>
        `AND(LEN(TRIM(${cellReference}))>0,ISNUMBER(SEARCH("@",${cellReference})),ISNUMBER(SEARCH(".",${cellReference},SEARCH("@",${cellReference})+2)),LEN(${cellReference})-LEN(SUBSTITUTE(${cellReference},"@",""))=1)`
    },
    validate: (value) => {
      const normalized = normalizeText(value);

      if (!normalized) {
        return false;
      }

      const parts = normalized.split("@");
      if (parts.length !== 2) {
        return false;
      }

      const [localPart, domainPart] = parts;
      if (!localPart || !domainPart) {
        return false;
      }

      const dotIndex = domainPart.indexOf(".");
      return dotIndex > 0 && dotIndex < domainPart.length - 1;
    }
  },
  Prefiks: {
    instruction:
      'Prefiks: This field is required and must start with "+", followed by digits only. "00" is not allowed and spaces are not allowed.',
    errorTitle: "Invalid Prefiks",
    errorMessage:
      'Prefiks is required and must start with "+", followed by digits only. "00" is not allowed and spaces are not allowed.',
    excel: {
      type: "custom",
      allowBlank: false,
      formula: (cellReference) => buildPrefixFormula(cellReference)
    },
    validate: (value) => /^\+(?!00)\d+$/.test(normalizeText(value))
  },
  Mobilnummer: {
    instruction: "Mobilnummer: This field is required and must contain exactly 8 digits with no spaces.",
    errorTitle: "Invalid phone number",
    errorMessage: "Mobilnummer is required and must contain exactly 8 digits with no spaces.",
    excel: {
      type: "custom",
      allowBlank: false,
      formula: (cellReference) =>
        `AND(LEN(${cellReference}&"")=8,ISNUMBER(1*${cellReference}),${cellReference}=INT(1*${cellReference}))`
    },
    validate: (value) => /^\d{8}$/.test(normalizeNumberLikeText(value))
  }
};

