export const WEIGHTED_LENGTH_LIMIT = 30;

// Han characters count as three; every other Unicode code point counts as one.
export function weightedLength(value) {
  return [...String(value ?? "")].reduce(
    (total, character) => total + (/\p{Script=Han}/u.test(character) ? 3 : 1),
    0,
  );
}

export function descriptionColumnIndexes(headers) {
  return headers
    .map((header, index) => ({ header: String(header ?? "").trim(), index }))
    .filter(({ header }) => /^(?:DESC[1-8]?|DESCRIPTION(?:\s*\d+)?)$/i.test(header));
}

export function evaluateDescriptionLength(headers, row) {
  const columns = descriptionColumnIndexes(headers);
  const values = columns.map(({ header, index }) => ({
    column: header,
    value: String(row[index] ?? "").trim(),
  }));
  const populated = values.filter(({ value }) => value !== "");
  const overLimit = populated
    .map(({ column, value }) => ({ column, value, weightedLength: weightedLength(value) }))
    .filter(({ weightedLength: length }) => length > WEIGHTED_LENGTH_LIMIT);
  const maxWeightedLength = populated.length
    ? Math.max(...populated.map(({ value }) => weightedLength(value)))
    : 0;

  return {
    descriptionColumns: columns.map(({ header }) => header),
    blankColumns: values.filter(({ value }) => value === "").map(({ column }) => column),
    overLimit,
    maxWeightedLength,
    status: overLimit.length ? "OVER_30" : (columns.length && !populated.length ? "BLANK" : "OK"),
  };
}
