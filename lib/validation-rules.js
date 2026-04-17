export const NAME_ALLOWED_CHARACTERS =
  "ABCDEFGHIJKLMNOPQRSTUVWXYZ\u00C6\u00D8\u00C5abcdefghijklmnopqrstuvwxyz\u00E6\u00F8\u00E5";

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

export const FIELD_RULES = {
  PersonID: {
    instruction: "PersonID: This field is required and must contain exactly 11 digits. Letters and other characters are not allowed.",
    errorTitle: "Invalid PersonID",
    errorMessage: "PersonID is required and must contain exactly 11 digits.",
    excel: {
      type: "custom",
      allowBlank: false,
      formula: (cellReference) =>
        `AND(LEN(${cellReference}&"")=11,ISNUMBER(1*${cellReference}),${cellReference}=INT(1*${cellReference}))`
    },
    validate: (value) => /^\d{11}$/.test(normalizeNumberLikeText(value))
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

