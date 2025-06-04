import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

export default function ExcelComparer() {
  const [files, setFiles] = useState({
    file1: null,
    file2: null,
    fileName1: '',
    fileName2: ''
  });
  const [data, setData] = useState({
    data1: [],
    data2: [],
    headers: []
  });
  const [diffs, setDiffs] = useState([]);
  const [stats, setStats] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [canCompare, setCanCompare] = useState(false);

  useEffect(() => {
    setCanCompare(data.data1.length > 0 && data.data2.length > 0);
  }, [data.data1, data.data2]);

  const handleFileUpload = (e, fileKey) => {
    const file = e.target.files[0];
    if (!file) return;

    setFiles(prev => ({
      ...prev,
      [fileKey]: file,
      [`fileName${fileKey === 'file1' ? '1' : '2'}`]: file.name
    }));

    const reader = new FileReader();

    reader.onload = (evt) => {
      try {
        const arrayBuffer = evt.target.result;
        const dataArr = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(dataArr, { type: "array" });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const sheetData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        setData(prev => ({
          ...prev,
          [fileKey === 'file1' ? 'data1' : 'data2']: sheetData,
          headers: sheetData.length > 0 ? sheetData[0] : prev.headers
        }));
      } catch (err) {
        setError(`Erreur lors de la lecture du fichier: ${err.message}`);
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

    setLoading(true);
    setError('');

    try {
      const { data1, data2 } = data;
      const maxRows = Math.max(data1.length, data2.length);
      const maxCols = Math.max(
        data1.length > 0 ? data1[0].length : 0,
        data2.length > 0 ? data2[0].length : 0
      );

      const differences = [];
      let diffCount = 0;

      for (let i = 0; i < maxRows; i++) {
        const row1 = data1[i] || Array(maxCols).fill('');
        const row2 = data2[i] || Array(maxCols).fill('');

        const diffRow = [];
        for (let j = 0; j < maxCols; j++) {
          const cell1 = row1[j] !== undefined ? row1[j] : '';
          const cell2 = row2[j] !== undefined ? row2[j] : '';

          if (cell1 !== cell2) {
            diffCount++;
            diffRow.push({
              value1: cell1,
              value2: cell2,
              isDiff: true
            });
          } else {
            diffRow.push({
              value: cell1,
              isDiff: false
            });
          }
        }
        differences.push(diffRow);
      }

      setDiffs(differences);
      setStats({
        totalCells: maxRows * maxCols,
        diffCells: diffCount,
        diffPercentage: (diffCount / (maxRows * maxCols) * 100).toFixed(2)
      });
    } catch (err) {
      setError(`Erreur lors de la comparaison: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const resetAll = () => {
    setFiles({
      file1: null,
      file2: null,
      fileName1: '',
      fileName2: ''
    });
    setData({
      data1: [],
      data2: [],
      headers: []
    });
    setDiffs([]);
    setStats(null);
    setError('');
  };

  return (
    <div className="p-4 max-w-6xl mx-auto">
      <h1 className="text-2xl font-bold mb-6 text-center">Comparateur de fichiers Excel</h1>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
        <div className="border p-4 rounded-lg">
          <label className="block mb-2 font-medium">Fichier 1</label>
          <div className="flex items-center gap-2">
            <label className="cursor-pointer bg-blue-500 hover:bg-blue-600 text-white px-4 py-2 rounded">
              Sélectionner
              <input
                type="file"
                className="hidden"
                accept=".xlsx,.xls,.csv"
                onChange={(e) => handleFileUpload(e, 'file1')}
              />
            </label>
            <span className="text-sm text-gray-600">{files.fileName1 || "Aucun fichier sélectionné"}</span>
          </div>
        </div>

        <div className="border p-4 rounded-lg">
          <label className="block mb-2 font-medium">Fichier 2</label>
          <div className="flex items-center gap-2">
            <label className="cursor-pointer bg-blue-500 hover:bg-blue-600 text-white px-4 py-2 rounded">
              Sélectionner
              <input
                type="file"
                className="hidden"
                accept=".xlsx,.xls,.csv"
                onChange={(e) => handleFileUpload(e, 'file2')}
              />
            </label>
            <span className="text-sm text-gray-600">{files.fileName2 || "Aucun fichier sélectionné"}</span>
          </div>
        </div>
      </div>

      <div className="flex justify-center gap-4 mb-6">
        <button
          onClick={compareFiles}
          disabled={!canCompare || loading}
          className={`px-6 py-2 rounded ${!canCompare || loading ? 'bg-gray-400 cursor-not-allowed' : 'bg-green-600 hover:bg-green-700 text-white'}`}
        >
          {loading ? 'Comparaison en cours...' : 'Comparer les fichiers'}
        </button>
        <button
          onClick={resetAll}
          className="px-6 py-2 bg-red-500 hover:bg-red-600 text-white rounded"
        >
          Réinitialiser
        </button>
      </div>

      {error && (
        <div className="mb-6 p-3 bg-red-100 border-l-4 border-red-500 text-red-700">
          <p>{error}</p>
        </div>
      )}

      {stats && (
        <div className="mb-6 p-4 bg-blue-50 rounded-lg">
          <h2 className="text-lg font-semibold mb-2">Statistiques de comparaison</h2>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="p-3 bg-white rounded shadow">
              <p className="text-sm text-gray-500">Cellules totales</p>
              <p className="text-xl font-bold">{stats.totalCells}</p>
            </div>
            <div className="p-3 bg-white rounded shadow">
              <p className="text-sm text-gray-500">Différences</p>
              <p className="text-xl font-bold text-red-600">{stats.diffCells}</p>
            </div>
            <div className="p-3 bg-white rounded shadow">
              <p className="text-sm text-gray-500">Pourcentage de différences</p>
              <p className="text-xl font-bold">{stats.diffPercentage}%</p>
            </div>
          </div>
        </div>
      )}

      {diffs.length > 0 && (
        <div className="overflow-x-auto">
          <h2 className="text-lg font-semibold mb-3">Résultats de la comparaison</h2>
          <div className="mb-2 text-sm text-gray-600">
            <span className="inline-block w-3 h-3 bg-red-200 mr-1"></span> Cellules différentes
          </div>
          <table className="w-full border-collapse">
            <thead>
              <tr>
                {data.headers.map((header, idx) => (
                  <th key={idx} className="border p-2 bg-gray-100">{header || `Colonne ${idx + 1}`}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {diffs.map((row, rowIndex) => (
                <tr key={rowIndex} className={row.some(cell => cell.isDiff) ? 'bg-red-50' : ''}>
                  {row.map((cell, colIndex) => (
                    <td
                      key={colIndex}
                      className={`border p-2 ${cell.isDiff ? 'bg-red-200' : ''}`}
                      title={cell.isDiff ? `${cell.value1} → ${cell.value2}` : ''}
                    >
                      {cell.isDiff ? (
                        <>
                          <span className="text-red-600 line-through">{cell.value1}</span>
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