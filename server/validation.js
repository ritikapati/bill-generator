const allowedSchemes = new Set(['General', 'SCP', 'TSP'])

export function validateRows(rows, { documentType } = {}) {
  if (!Array.isArray(rows) || rows.length === 0) {
    throw new Error('Add at least one chart row before generating a bill.')
  }

  return rows.map((row, index) => {
    const rowNumber = index + 1
    const ada = typeof row.ada === 'string' ? row.ada.trim() : ''
    const rsk = typeof row.rsk === 'string' ? row.rsk.trim() : ''
    const crop = typeof row.crop === 'string' ? row.crop.trim() : ''
    const scheme = typeof row.scheme === 'string' ? row.scheme.trim() : ''
    const quantity = row.quantity

    if (!ada) throw new Error(`Row ${rowNumber}: ADA is required.`)
    if (!allowedSchemes.has(scheme)) throw new Error(`Row ${rowNumber}: scheme must be General, SCP or TSP.`)
    if (!crop) throw new Error(`Row ${rowNumber}: crop or product is required.`)
    if (documentType === 'challan' && !rsk) throw new Error(`Row ${rowNumber}: RSK is required for a Delivery Challan.`)
    if (quantity === null || quantity === '' || !Number.isFinite(Number(quantity)) || Number(quantity) < 0) {
      throw new Error(`Row ${rowNumber}: enter a valid non-negative quantity.`)
    }

    return { ada, rsk, scheme, crop, quantity: Number(quantity) }
  })
}