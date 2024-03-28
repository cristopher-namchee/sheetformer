import * as excel from "exceljs";

import { OCRField, OCRRead, OCRResponse } from "../../types";
import { Keyword } from "./types";

const keywordNonTableMapping: Keyword = {
  JENIS_PENGAJUAN: { field: "debit_account_request_type" },
  NAMA_LENGKAP: { field: "complete_name" },
  NO_KTP: { field: "id_card_number" },
  TANGGAL_MASA_BERLAKU: {
    field: "validity_period",
    format: "date",
    segment: "date",
    expected_length: 2,
  },
  BULAN_MASA_BERLAKU: {
    field: "validity_period",
    format: "date",
    segment: "month",
    expected_length: 2,
  },
  TAHUN_MASA_BERLAKU: {
    field: "validity_period",
    format: "date",
    segment: "year",
    expected_length: 2,
  },
  BERTINDAK_SEBAGAI: { field: "act_for" },
  BERTINDAK_SEBAGAI_NASABAH_BADAN_JABATAN: { field: "position" },
  BERTINDAK_SEBAGAI_NASABAH_BADAN_NAMA_PT: { field: "company_name" },
  ALAMAT: { field: "address" },
  RT: {
    field: "rt_rw",
    format: "string",
    delimiter: "/",
    segment: 0,
    padding_char: "0",
    expected_length: 3,
  },
  RW: {
    field: "rt_rw",
    format: "string",
    delimiter: "/",
    segment: 1,
    padding_char: "0",
    expected_length: 3,
  },
  KELURAHAN: { field: "district" },
  KOTA: { field: "city" },
  KODE_POS: { field: "zip_code" },
  PROVINSI: { field: "province" },
  KUASA_UNTUK_NAMA: { field: "attorney_in_fact_name" },
  KUASA_UNTUK_ALAMAT: { field: "attorney_in_fact_address" },
  KUASA_UNTUK_NO_KTP_SIM_PASPOR: { field: "attorney_in_fact_id_card_number" },
  KUASA_UNTUK_KODE_PERUSAHAAN: { field: "attorney_in_fact_company_code" },
  KODE_AREA_TELEPON_KANTOR: {
    field: "office_phone",
    format: "string",
    delimiter: /[ \-]/,
    segment: 0,
  },
  TELEPON_KANTOR: {
    field: "office_phone",
    format: "string",
    delimiter: /[ \-]/,
    segment: 1,
  },
  KODE_AREA_TELEPON_RUMAH: {
    field: "home_phone",
    format: "string",
    delimiter: /[ \-]/,
    segment: 0,
  },
  TELEPON_RUMAH: {
    field: "home_phone",
    format: "string",
    delimiter: /[ \-]/,
    segment: 1,
  },
  NO_HP_1: { field: "mobile_phone_number_1" },
  NO_HP_2: { field: "mobile_phone_number_2" },
  NOMOR_REKENING: { field: "bank_account_number" },
  NAMA_BANK: { field: "bank_name" },
  NAMA_PEMILIK_REKENING: { field: "bank_account_owner_name" },
  JENIS_REKENING: { field: "account_type" },
  EMAIL: { field: "email_address" },
  MATA_UANG_REKENING: { field: "account_currency" },
  MATA_UANG_REKENING_LAINNYA: { field: "account_other_currency" },
  MATA_UANG_POLIS: { field: "policy_currency" },
  MATA_UANG_POLIS_LAINNYA: { field: "policy_other_currency" },
  NO_SPA_POLIS_INVOICE: { field: "policy_number" },
  NAMA_PEMEGANG_POLIS: { field: "policyholder_name" },
  HUB_DG_PEMEGANG_POLIS: { field: "relationship_with_policyholder" },
  HUB_DG_PEMEGANG_POLIS_LAINNYA: {
    field: "other_relationship_with_policyholder",
  },
  TEMPAT_PENANDATANGANAN: { field: "signing_location" },
  TANGGAL_PENANDATANGANAN: {
    field: "signing_date",
    format: "date",
    segment: "date",
    expected_length: 2,
  },
  BULAN_PENANDATANGANAN: {
    field: "signing_date",
    format: "date",
    segment: "month",
    expected_length: 2,
  },
  TAHUN_PENANDATANGANAN: {
    field: "signing_date",
    format: "date",
    segment: "year",
    expected_length: 4,
  },
  NAMA_PEMBERI_KUASA_AUTOFILLED: { field: "complete_name" },
  NAMA_PEMEGANG_POLIS_AUTOFILLED: { field: "policyholder_name" },
};

const keywordTableAccountMapping: Keyword = {
  NOMOR_REKENING: { field: "bank_account_number" },
  NAMA_PEMILIK_REKENING: { field: "bank_account_owner_name" },
};

const keywordTablePolicyMapping: Keyword = {
  NOMOR_POLIS: { field: "policy_number" },
  NAMA_PEMEGANG_POLIS: { field: "policyholder_name" },
  HUB_PEMILIK_REKENING_DAN_PEMEGANG_POLIS: {
    field: "other_relationship_with_policyholder",
  },
};

function formatAsDate(
  ocrText: string,
  segment: string | null = null,
  expected_length: number | null = null
): string | null {
  const dateSplit = ocrText.split("-");
  let dateParts = [];

  if (!isNaN(Date.parse(ocrText))) {
    const date = new Date(ocrText);
    dateParts[0] =
      (dateSplit.length === 3 && dateSplit[2].length === 2) ||
      (dateSplit.length === 2 &&
        dateSplit[0].length == 2 &&
        dateSplit[1].length === 2)
        ? date.getDate().toString()
        : null;
    dateParts[1] =
      (dateSplit.length === 3 && dateSplit[1].length === 2) ||
      (dateSplit.length === 2 &&
        dateSplit[0].length == 4 &&
        dateSplit[1].length == 2) ||
      (dateSplit.length === 2 && dateSplit[0].length == 2)
        ? (date.getMonth() + 1).toString()
        : null;
    dateParts[2] =
      dateSplit.length >= 1 && dateSplit[0].length === 4
        ? date.getFullYear().toString()
        : null;
  } else {
    dateParts = dateSplit;
  }

  for (let idx = 0; idx < dateParts.length; idx++) {
    dateParts[idx] = String(parseInt(dateParts[idx], 10)).padStart(2, "0");
  }

  const returnMapper = {
    date: dateParts?.[0],
    month: dateParts?.[1],
    year:
      !expected_length || expected_length >= 4 || !dateParts?.[2]
        ? dateParts?.[2]
        : dateParts?.[2].slice(-expected_length),
    full: dateParts.filter(Boolean).join("/"),
  };

  const segmentKey = segment ?? "full";
  const returnValue =
    segmentKey in returnMapper ? returnMapper[segmentKey] : null;

  return [undefined, "NaN"].includes(returnValue) ? null : returnValue;
}

function formatAsString(
  ocrText: string,
  delimiter?: string,
  segment?: number,
  expected_length?: number,
  padding_char?: string
): string {
  if (delimiter !== undefined && segment !== undefined) {
    const parts = ocrText.split(delimiter);

    ocrText =
      parts.length && parts[segment]
        ? padString(parts[segment], expected_length, padding_char)
        : "";
  }

  return ocrText;
}

function padString(
  ocrText: string,
  expected_length?: number,
  padding_char?: string
): string {
  if (expected_length && padding_char) {
    if (ocrText.length >= expected_length) {
      return ocrText.slice(-expected_length);
    } else {
      const paddingCount = expected_length - ocrText.length;
      const padding = padding_char.repeat(paddingCount);
      return padding + ocrText;
    }
  }

  return ocrText;
}

function formatFieldOCR(ocrFieldObject, mappingData): string {
  const { value } = ocrFieldObject ?? {};

  const { format, delimiter, segment, expected_length, padding_char } =
    mappingData ?? {};

  let ocrText = [null, undefined, ""].includes(value) ? null : value;
  if (ocrText !== null && format === "number") {
    ocrText = parseFloat(ocrText).toFixed(2);
  } else if (ocrText !== null && format === "date") {
    ocrText = formatAsDate(ocrText, segment, expected_length);
  } else if (ocrText !== null && format === "string") {
    ocrText = formatAsString(
      ocrText,
      delimiter,
      segment,
      expected_length,
      padding_char
    );
  }

  return ocrText;
}

function fillOCRTableDataToSheet(
  accountDetails: OCRField[],
  policyDetails: OCRField[]
): [string[][], number, number] {
  const tables: string[][] = [
    [
      "NOMOR_REKENING",
      "NAMA_PEMILIK_REKENING",
      "",
      "NOMOR_POLIS",
      "NAMA_PEMEGANG_POLIS",
      "HUB_PEMILIK_REKENING_DAN_PEMEGANG_POLIS",
    ],
  ];

  accountDetails = accountDetails ?? [];
  policyDetails = policyDetails ?? [];

  let idx = 0;

  while (idx < accountDetails.length || idx < policyDetails.length) {
    const row: string[] = [];
    if (idx < accountDetails.length) {
      const detail = accountDetails[idx];

      for (const fieldInfo of Object.values(keywordTableAccountMapping)) {
        const value = formatFieldOCR(detail[fieldInfo.field], fieldInfo);

        row.push(value);
      }
    }

    // pad the spaces
    while (row.length < 3) {
      row.push("");
    }

    if (idx < policyDetails.length) {
      const detail = policyDetails[idx];
      for (const fieldInfo of Object.values(keywordTablePolicyMapping)) {
        const value = formatFieldOCR(detail[fieldInfo.field], fieldInfo);

        row.push(value);
      }
    }

    tables.push(row);

    idx++;
  }

  return [tables, accountDetails.length, policyDetails.length];
}

function beautifySheet(sheet: excel.Worksheet) {
  const baseFontStyle = {
    name: "Times New Roman",
    size: 11,
  };
  for (let row = 1; row <= 49; row++) {
    sheet.getCell(`A${row}`).style = {
      font: {
        ...baseFontStyle,
        bold: true,
      },
    };
  }
}

function applyBorder(
  sheet: excel.Worksheet,
  r: { start: number; end: number },
  c: { start: number; end: number }
) {
  const borderStyle: excel.Border = {
    style: "thin",
    color: {
      argb: "FF000000",
    },
  };
  const border: Partial<excel.Borders> = {
    top: borderStyle,
    bottom: borderStyle,
    left: borderStyle,
    right: borderStyle,
  };

  for (let col = c.start; col <= c.end; col++) {
    for (let row = r.start; row <= r.end; row++) {
      const cell = sheet.getCell(row, col);
      cell.style = {
        ...cell.style,
        border,
      };
    }
  }

  for (let col = c.start; col <= c.end; col++) {
    const cell = sheet.getCell(r.start, col);

    cell.style = {
      ...cell.style,
      font: {
        ...cell.font,
        bold: true,
      },
    };
  }
}

function beautifyColumn(sheet: excel.Worksheet) {
  sheet.columns.forEach((column) => {
    if (column.values) {
      const lengths = column.values.map((v) => (v ? v.toString().length : 0));
      const maxLength = Math.max(
        ...lengths.filter((v) => typeof v === "number")
      );

      column.width = maxLength + 12;
    }
  });
}

function handleLifetimeValidityPeriod(rows: string[][], read: OCRRead) {
  const validityDate = rows.findIndex(
    (row) => row[0] === "TANGGAL_MASA_BERLAKU"
  );
  const validityMonth = rows.findIndex(
    (row) => row[0] === "BULAN_MASA_BERLAKU"
  );
  const validityYear = rows.findIndex((row) => row[0] === "TAHUN_MASA_BERLAKU");
  const period = read.validity_period
    ? ((read.validity_period as OCRField).value as string)
    : "";

  if (period.toLowerCase() === "seumur hidup") {
    rows[validityDate][1] = "Seumur Hidup";
    rows[validityMonth][1] = "";
    rows[validityYear][1] = "";
  }
}

export default async function exportToSheet(
  response: OCRResponse,
  documentName: string
): Promise<excel.Buffer> {
  const workbook = new excel.Workbook();

  // truncate limit
  documentName = documentName.slice(0, 31);

  const sheet = workbook.addWorksheet(documentName);
  const rows = [["Extraction Result"], [""], ["File Name", documentName], [""]];

  const read = response.read as OCRRead;
  if (read) {
    for (const [key, fieldInfo] of Object.entries(keywordNonTableMapping)) {
      const value = formatFieldOCR(read[fieldInfo.field], fieldInfo);
      rows.push([key, value]);
    }

    const [table, accounts, policies] = fillOCRTableDataToSheet(
      read.bank_account_details as OCRField[],
      read.policy_details as OCRField[]
    );

    if (table) {
      rows.push([""], ...table);
    }

    handleLifetimeValidityPeriod(rows, read);

    sheet.addRows(rows);
    beautifySheet(sheet);

    applyBorder(
      sheet,
      {
        start: 51,
        end: 51 + accounts,
      },
      {
        start: 1,
        end: 2,
      }
    );

    applyBorder(
      sheet,
      {
        start: 51,
        end: 51 + policies,
      },
      {
        start: 4,
        end: 6,
      }
    );

    if (accounts) {
      sheet.getCell("B31").value = "";
      sheet.getCell("B33").value = "";
    }

    beautifyColumn(sheet);
  }

  return workbook.xlsx.writeBuffer();
}
