const schemeFields = [
  ['General', 'general'],
  ['SCP', 'scp'],
  ['TSP', 'tsp'],
]

export function expandChartGroups(groups) {
  if (!Array.isArray(groups)) throw new Error('No chart groups were returned by the model.')

  const rows = []
  groups.forEach((group) => {
    for (const [scheme, field] of schemeFields) {
      const quantity = group.quantities?.[field]
      if (quantity === null || quantity === undefined) continue
      if (!Number.isFinite(quantity) || quantity < 0) continue
      rows.push({
        ada: typeof group.ada === 'string' ? group.ada.trim() : '',
        rsk: typeof group.rsk === 'string' ? group.rsk.trim() : '',
        scheme,
        crop: typeof group.crop === 'string' ? group.crop.trim() : '',
        quantity,
      })
    }
  })

  return rows
}