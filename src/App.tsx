import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

type ExcelData = (string | number | boolean | null)[][];
type DiffCell =
  | {
      value1: string | number | boolean | null;
      value2: string | number | boolean | null;
      isDiff: true;
    }
  | { value: string | number | boolean | null; isDiff: false };

type Stats = {
  totalCells: number;
  diffCells: number;
  diffPercentage: string;
};

export default function ExcelComparer() {
  const [files, setFiles] = useState({
    file1: null as File | null,
    file2: null as File | null,
    fileName1: "",
    fileName2: "",
  });

  const [data, setData] = useState<{
    data1: ExcelData;
    data2: ExcelData;
    headers: (string | number)[];
  }>({
    data1: [],
    data2: [],
    headers: [],
  });

  const [diffs, setDiffs] = useState<DiffCell[][]>([]);
  const [stats, setStats] = useState<Stats | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [canCompare, setCanCompare] = useState(false);

  const isValidFileType = (fileName: string): boolean => {
    const validExtensions = [".xlsx", ".xls", ".csv"];
    const extension = fileName.slice(fileName.lastIndexOf(".")).toLowerCase();
    return validExtensions.includes(extension);
  };

  useEffect(() => {
    setCanCompare(data.data1.length > 0 && data.data2.length > 0);
  }, [data.data1, data.data2]);

  const handleFileUpload = (
    e: React.ChangeEvent<HTMLInputElement>,
    fileKey: "file1" | "file2",
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!isValidFileType(file.name)) {
      setError(
        `Type de fichier non valide. Veuillez sélectionner un fichier .xlsx, .xls ou .csv`,
      );
      return;
    }

    setFiles((prev) => ({
      ...prev,
      [fileKey]: file,
      [`fileName${fileKey === "file1" ? "1" : "2"}`]: file.name,
    }));

    const reader = new FileReader();

    reader.onload = (evt) => {
      try {
        const arrayBuffer = (evt.target?.result ??
          new ArrayBuffer(0)) as ArrayBuffer;
        const dataArr = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(dataArr, { type: "array" });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const sheetData = XLSX.utils.sheet_to_json(firstSheet, {
          header: 1,
        }) as ExcelData;

        setData((prev) => ({
          ...prev,
          [fileKey === "file1" ? "data1" : "data2"]: sheetData,
          headers:
            sheetData.length > 0
              ? sheetData[0].filter(
                  (cell): cell is string | number =>
                    typeof cell === "string" || typeof cell === "number",
                )
              : prev.headers,
        }));
      } catch (err) {
        setError(
          `Erreur lors de la lecture du fichier: ${(err as Error).message}`,
        );
      }
    };

    reader.onerror = () => {
      setError("Erreur lors de la lecture du fichier");
    };

    reader.readAsArrayBuffer(file);
  };

  const compareFiles = () => {
    if (!files.file1 || !files.file2) {
      setError("Veuillez sélectionner les deux fichiers");
      return;
    }

    if (
      !isValidFileType(files.fileName1) ||
      !isValidFileType(files.fileName2)
    ) {
      setError("Un ou plusieurs fichiers ont un format non valide");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const { data1, data2 } = data;
      const maxRows = Math.max(data1.length, data2.length);
      const maxCols = Math.max(
        data1.length > 0 ? data1[0].length : 0,
        data2.length > 0 ? data2[0].length : 0,
      );

      const differences: DiffCell[][] = [];
      let diffCount = 0;

      for (let i = 0; i < maxRows; i++) {
        const row1 = data1[i] || Array(maxCols).fill("");
        const row2 = data2[i] || Array(maxCols).fill("");

        const diffRow: DiffCell[] = [];

        for (let j = 0; j < maxCols; j++) {
          const cell1 = row1[j] ?? "";
          const cell2 = row2[j] ?? "";

          if (cell1 !== cell2) {
            diffCount++;
            diffRow.push({
              value1: cell1,
              value2: cell2,
              isDiff: true,
            });
          } else {
            diffRow.push({
              value: cell1,
              isDiff: false,
            });
          }
        }

        differences.push(diffRow);
      }

      setDiffs(differences);
      setStats({
        totalCells: maxRows * maxCols,
        diffCells: diffCount,
        diffPercentage: ((diffCount / (maxRows * maxCols)) * 100).toFixed(2),
      });
    } catch (err) {
      setError(`Erreur lors de la comparaison: ${(err as Error).message}`);
    } finally {
      setLoading(false);
    }
  };

  const resetAll = () => {
    setFiles({
      file1: null,
      file2: null,
      fileName1: "",
      fileName2: "",
    });
    setData({
      data1: [],
      data2: [],
      headers: [],
    });
    setDiffs([]);
    setStats(null);
    setError("");
  };

  return (
    <div className="mx-auto max-w-6xl p-4">
      <h1 className="mb-6 text-center text-2xl font-bold">
        Comparateur de fichiers Excel
      </h1>

      {/* Upload section */}
      <div className="mb-6 grid grid-cols-1 gap-6 md:grid-cols-2">
        {["file1", "file2"].map((fileKey, index) => (
          <div key={fileKey} className="rounded-lg border p-4">
            <label className="mb-2 block font-medium">
              Fichier {index + 1}
            </label>
            <div className="flex items-center gap-2">
              <label className="cursor-pointer rounded bg-blue-500 px-4 py-2 text-white hover:bg-blue-600">
                Sélectionner
                <input
                  type="file"
                  className="hidden"
                  accept=".xlsx,.xls,.csv"
                  onChange={(e) =>
                    handleFileUpload(e, fileKey as "file1" | "file2")
                  }
                />
              </label>
              <span className="text-sm text-gray-600">
                {files[`fileName${index + 1}` as "fileName1" | "fileName2"] ||
                  "Aucun fichier sélectionné"}
              </span>
            </div>
          </div>
        ))}
      </div>

      {/* Buttons */}
      <div className="mb-6 flex justify-center gap-4">
        <button
          onClick={compareFiles}
          disabled={!canCompare || loading}
          className={`rounded px-6 py-2 ${!canCompare || loading ? "cursor-not-allowed bg-gray-400" : "bg-green-600 text-white hover:bg-green-700"}`}
        >
          {loading ? "Comparaison en cours..." : "Comparer les fichiers"}
        </button>
        <button
          onClick={resetAll}
          className="rounded bg-red-500 px-6 py-2 text-white hover:bg-red-600"
        >
          Réinitialiser
        </button>
      </div>

      {/* Errors */}
      {error && (
        <div className="mb-6 border-l-4 border-red-500 bg-red-100 p-3 text-red-700">
          <p>{error}</p>
        </div>
      )}

      {/* Stats */}
      {stats && (
        <div className="mb-6 rounded-lg bg-blue-50 p-4">
          <h2 className="mb-2 text-lg font-semibold">
            Statistiques de comparaison
          </h2>
          <div className="grid grid-cols-1 gap-4 md:grid-cols-3">
            <div className="rounded bg-white p-3 shadow">
              <p className="text-sm text-gray-500">Cellules totales</p>
              <p className="text-xl font-bold">{stats.totalCells}</p>
            </div>
            <div className="rounded bg-white p-3 shadow">
              <p className="text-sm text-gray-500">Différences</p>
              <p className="text-xl font-bold text-red-600">
                {stats.diffCells}
              </p>
            </div>
            <div className="rounded bg-white p-3 shadow">
              <p className="text-sm text-gray-500">
                Pourcentage de différences
              </p>
              <p className="text-xl font-bold">{stats.diffPercentage}%</p>
            </div>
          </div>
        </div>
      )}

      {/* Diff Table */}
      {diffs.length > 0 && (
        <div className="overflow-x-auto">
          <h2 className="mb-3 text-lg font-semibold">
            Résultats de la comparaison
          </h2>
          <div className="mb-2 text-sm text-gray-600">
            <span className="mr-1 inline-block h-3 w-3 bg-red-200"></span>{" "}
            Cellules différentes
          </div>
          <table className="w-full border-collapse">
            <thead>
              <tr>
                {data.headers.map((header, idx) => (
                  <th key={idx} className="border bg-gray-100 p-2">
                    {header || `Colonne ${idx + 1}`}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {diffs.map((row, rowIndex) => (
                <tr
                  key={rowIndex}
                  className={row.some((cell) => cell.isDiff) ? "bg-red-50" : ""}
                >
                  {row.map((cell, colIndex) => (
                    <td
                      key={colIndex}
                      className={`border p-2 ${cell.isDiff ? "bg-red-200" : ""}`}
                      title={
                        cell.isDiff ? `${cell.value1} → ${cell.value2}` : ""
                      }
                    >
                      {cell.isDiff ? (
                        <>
                          <span className="text-red-600 line-through">
                            {cell.value1}
                          </span>
                          <br />
                          <span className="text-green-600">{cell.value2}</span>
                        </>
                      ) : (
                        cell.value
                      )}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
