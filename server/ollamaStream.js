export async function readOllamaStream(stream) {
  const reader = stream.getReader()
  const decoder = new TextDecoder()
  let buffered = ''
  let content = ''
  let finalChunk = null

  function consumeLine(line) {
    if (!line.trim()) return
    const chunk = JSON.parse(line)
    if (chunk.error) throw new Error(chunk.error)
    content += chunk.message?.content || ''
    if (chunk.done) finalChunk = chunk
  }

  try {
    while (true) {
      const { done, value } = await reader.read()
      if (done) break
      buffered += decoder.decode(value, { stream: true })
      const lines = buffered.split('\n')
      buffered = lines.pop() || ''
      lines.forEach(consumeLine)
    }

    buffered += decoder.decode()
    consumeLine(buffered)
    return { content, finalChunk }
  } catch (error) {
    await reader.cancel().catch(() => {})
    throw error
  }
}
