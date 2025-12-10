export interface DocumentPayload {
  title: string
  problemStatement: string
  objectives: string[]
  requirements: string[]
  manualSteps: string[]
  automationIdeas: string
}

export async function generateDocx(payload: DocumentPayload) {
  try {
    const response = await fetch("/api/generate-docx", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    })

    if (!response.ok) {
      const error = await response.json()
      throw new Error(error.error || "Failed to generate document")
    }

    const data = await response.json()
    return {
      fileName: data.fileName,
      fileBase64: data.fileBase64,
    }
  } catch (error) {
    console.error("Error calling docx API:", error)
    throw error
  }
}

// Helper to download file in browser
export function downloadFile(fileName: string, fileBase64: string) {
  const binaryString = atob(fileBase64)
  const bytes = new Uint8Array(binaryString.length)
  for (let i = 0; i < binaryString.length; i++) {
    bytes[i] = binaryString.charCodeAt(i)
  }

  const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" })
  const url = window.URL.createObjectURL(blob)
  const a = document.createElement("a")
  a.href = url
  a.download = fileName
  a.click()
  window.URL.revokeObjectURL(url)
}
