import sharp from 'sharp'

export async function prepareChartImage(buffer) {
  return sharp(buffer)
    .rotate()
    // Keep the source's detail without enlarging it. Large vision inputs make
    // local CPU inference stall, so cap full-page charts and let users crop when needed.
    .resize({ width: 1024, height: 1280, fit: 'inside', withoutEnlargement: true })
    .flatten({ background: '#ffffff' })
    .sharpen({ sigma: 0.8 })
    .jpeg({ quality: 94 })
    .toBuffer()
}
