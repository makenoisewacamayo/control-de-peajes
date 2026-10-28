"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import Header from "../../components/header";

const FLEET = [
  "HPLN-001",
  "HPLN-002",
  "HPLN-003",
  "HPLN-004",
  "HPLN-005",
  "HPLN-006",
  "HPLN-007",
  "HPLN-008",
  "HPLN-009",
  "HPLN-010",
];

type Charge = {
  patente: string;
  fecha: string;
  autopista: string;
  portico: string;
  montoCobrado: number;
  tarifaEsperada: number;
  rebaja: number;
  perteneceFlota: boolean;
};

const money = new Intl.NumberFormat("es-CL", {
  style: "currency",
  currency: "CLP",
  maximumFractionDigits: 0,
});

function normalize(value: unknown) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_|_$/g, "");
}

function parseAmount(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value === null || value === undefined || value === "") return null;

  const digits = String(value).replace(/[^\d-]/g, "");
  if (!digits || digits === "-") return null;

  const amount = Number(digits);
  return Number.isFinite(amount) ? amount : null;
}

function getValue(
  row: Record<string, unknown>,
  names: string[],
): unknown {
  for (const name of names) {
    const key = Object.keys(row).find((item) => normalize(item) === name);
    if (key) return row[key];
  }
  return undefined;
}

function downloadWorkbook(rows: object[], filename: string) {
  const sheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "Cobros");
  XLSX.writeFile(workbook, filename);
}

export default function Home() {
  const [charges, setCharges] = useState<Charge[]>([]);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");

  const totalRebaja = charges.reduce(
    (sum, charge) => sum + charge.rebaja,
    0,
  );
  const cobrosRevisados = charges.filter(
    (charge) => charge.perteneceFlota,
  ).length;

  async function handleFile(file?: File) {
    if (!file) return;

    setError("");
    setCharges([]);

    try {
      const bytes = await file.arrayBuffer();
      const workbook = XLSX.read(bytes);
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

      if (!firstSheet) throw new Error("El archivo no contiene hojas.");

      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(
        firstSheet,
        { defval: "" },
      );

      if (rows.length === 0) {
        throw new Error("La primera hoja está vacía.");
      }

      const result: Charge[] = rows.map((row, index) => {
        const patente = String(
          getValue(row, ["patente", "placa"]) ?? "",
        )
          .trim()
          .toUpperCase();

        const fecha = String(getValue(row, ["fecha"]) ?? "");
        const autopista = String(
          getValue(row, ["autopista", "concesionaria"]) ?? "",
        );
        const portico = String(
          getValue(row, ["portico", "pórtico"]) ?? "",
        );

        const montoCobrado = parseAmount(
          getValue(row, ["monto_cobrado", "cobrado"]),
        );
        const tarifaEsperada = parseAmount(
          getValue(row, ["tarifa_esperada", "monto_esperado", "tarifa"]),
        );

        if (!patente || montoCobrado === null || tarifaEsperada === null) {
          throw new Error(
            `Revisa la fila ${index + 2}: se requieren patente, monto_cobrado y tarifa_esperada.`,
          );
        }

        const perteneceFlota = FLEET.includes(patente);
        const rebaja = perteneceFlota
          ? Math.max(0, montoCobrado - tarifaEsperada)
          : 0;

        return {
          patente,
          fecha,
          autopista,
          portico,
          montoCobrado,
          tarifaEsperada,
          rebaja,
          perteneceFlota,
        };
      });

      setCharges(result);
      setFileName(file.name);
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "No se pudo leer el archivo.",
      );
    }
  }

  function downloadTemplate() {
    downloadWorkbook(
      [
        {
          patente: "HPLN-001",
          fecha: "2026-10-01",
          autopista: "Autopista de ejemplo",
          portico: "Pórtico 1",
          monto_cobrado: 400,
          tarifa_esperada: 200,
        },
      ],
      "plantilla_cobros_peajes.xlsx",
    );
  }

  function downloadRebajas() {
    const fleetCharges = charges
      .filter((charge) => charge.perteneceFlota)
      .map((charge) => ({
        patente: charge.patente,
        fecha: charge.fecha,
        autopista: charge.autopista,
        portico: charge.portico,
        monto_cobrado: charge.montoCobrado,
        tarifa_esperada: charge.tarifaEsperada,
        rebaja_sugerida: charge.rebaja,
      }));

    downloadWorkbook(fleetCharges, "rebajas_peajes.xlsx");
  }

  return (
    <main className="min-h-screen bg-slate-50 p-6 text-slate-900">
      <div className="mx-auto max-w-6xl space-y-6">
        <Header />

        <section className="grid gap-4 sm:grid-cols-3">
          <div className="rounded-xl bg-white p-5 shadow-sm">
            <p className="text-sm text-slate-500">Buses registrados</p>
            <p className="mt-1 text-2xl font-bold">{FLEET.length}</p>
          </div>
          <div className="rounded-xl bg-white p-5 shadow-sm">
            <p className="text-sm text-slate-500">Cobros de la flota revisados</p>
            <p className="mt-1 text-2xl font-bold">{cobrosRevisados}</p>
          </div>
          <div className="rounded-xl bg-white p-5 shadow-sm">
            <p className="text-sm text-slate-500">Rebaja sugerida total</p>
            <p className="mt-1 text-2xl font-bold text-emerald-700">
              {money.format(totalRebaja)}
            </p>
          </div>
        </section>

        <section className="rounded-xl bg-white p-5 shadow-sm">
          <h2 className="text-lg font-semibold">Cargar cobros</h2>
          <p className="mt-1 text-sm text-slate-600">
            El Excel debe incluir: patente, fecha, autopista, portico,
            monto_cobrado y tarifa_esperada.
          </p>

          <div className="mt-4 flex flex-wrap gap-3">
            <label className="cursor-pointer rounded-lg bg-blue-700 px-4 py-2 font-medium text-white hover:bg-blue-800">
              Seleccionar Excel
              <input
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={(event) => handleFile(event.target.files?.[0])}
              />
            </label>

            <button
              onClick={downloadTemplate}
              className="rounded-lg border border-slate-300 px-4 py-2 font-medium hover:bg-slate-50"
            >
              Descargar plantilla de ejemplo
            </button>

            {charges.length > 0 && (
              <button
                onClick={downloadRebajas}
                className="rounded-lg bg-emerald-700 px-4 py-2 font-medium text-white hover:bg-emerald-800"
              >
                Descargar Excel de rebajas
              </button>
            )}
          </div>

          {fileName && (
            <p className="mt-3 text-sm text-slate-500">
              Archivo cargado: {fileName}
            </p>
          )}
          {error && (
            <p className="mt-3 text-sm font-medium text-red-700">{error}</p>
          )}
        </section>

        <section className="rounded-xl bg-white p-5 shadow-sm">
          <h2 className="text-lg font-semibold">Detalle de cobros</h2>
          <div className="mt-4 overflow-x-auto">
            <table className="w-full min-w-[850px] text-left text-sm">
              <thead className="border-b text-slate-500">
                <tr>
                  <th className="p-3">Patente</th>
                  <th className="p-3">Fecha</th>
                  <th className="p-3">Autopista / pórtico</th>
                  <th className="p-3">Cobrado</th>
                  <th className="p-3">Esperado</th>
                  <th className="p-3">Rebaja sugerida</th>
                  <th className="p-3">Estado</th>
                </tr>
              </thead>
              <tbody>
                {charges.map((charge, index) => (
                  <tr key={`${charge.patente}-${index}`} className="border-b">
                    <td className="p-3 font-medium">{charge.patente}</td>
                    <td className="p-3">{charge.fecha}</td>
                    <td className="p-3">
                      {[charge.autopista, charge.portico]
                        .filter(Boolean)
                        .join(" · ") || "—"}
                    </td>
                    <td className="p-3">{money.format(charge.montoCobrado)}</td>
                    <td className="p-3">{money.format(charge.tarifaEsperada)}</td>
                    <td className="p-3 font-semibold text-emerald-700">
                      {money.format(charge.rebaja)}
                    </td>
                    <td className="p-3">
                      {charge.perteneceFlota
                        ? charge.rebaja > 0
                          ? "Revisar cobro"
                          : "Sin diferencia"
                        : "Patente fuera de la flota demo"}
                    </td>
                  </tr>
                ))}

                {charges.length === 0 && (
                  <tr>
                    <td colSpan={7} className="p-6 text-center text-slate-500">
                      Carga un Excel para ver los cobros.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    </main>
  );
}