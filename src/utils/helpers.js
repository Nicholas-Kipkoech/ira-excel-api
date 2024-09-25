const formatOracleData = (fetchedData) => {
  const { metaData, rows } = fetchedData;

  // Map the rows to a more meaningful format based on metaData
  return rows.map((row) => {
    return metaData.reduce((formattedRow, meta, index) => {
      formattedRow[meta.name] = row[index];
      return formattedRow;
    }, {});
  });
};
export default formatOracleData;
