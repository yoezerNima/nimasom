"use client"

import { useState } from "react"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { generateDocx, downloadFile } from "@/lib/docx-client"

export default function Home() {
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const handleGenerateExample = async () => {
    setLoading(true)
    setError(null)

    try {
      const result = await generateDocx({
        title: "Automation Project Proposal",
        problemStatement:
          "Our current process requires manual data entry from multiple sources, which is time-consuming and error-prone.",
        objectives: [
          "Reduce manual data entry by 80%",
          "Improve data accuracy and consistency",
          "Free up team members for higher-value tasks",
        ],
        requirements: [
          "Integration with existing CRM system",
          "Real-time data synchronization",
          "Error logging and notification system",
        ],
        manualSteps: [
          "Extract data from source system",
          "Validate data format and completeness",
          "Enter data into destination system",
          "Verify entries and reconcile discrepancies",
        ],
        automationIdeas:
          "Use Power Automate to automatically pull data from source systems, validate using business rules, and sync to destination platforms. Implement exception handling for edge cases.",
      })

      downloadFile(result.fileName, result.fileBase64)
    } catch (err) {
      setError(err instanceof Error ? err.message : "An error occurred")
    } finally {
      setLoading(false)
    }
  }

  return (
    <main className="flex min-h-screen items-center justify-center bg-background p-4">
      <Card className="w-full max-w-md">
        <CardHeader>
          <CardTitle>DOCX Generation Service</CardTitle>
          <CardDescription>Generate Word documents for Power Automate integration</CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          <p className="text-sm text-muted-foreground">
            This service generates .docx files that can be called directly from Power Automate flows. Click the button
            below to test with example data.
          </p>

          {error && (
            <div className="rounded-md border border-destructive bg-destructive/10 p-3 text-sm text-destructive">
              {error}
            </div>
          )}

          <Button onClick={handleGenerateExample} disabled={loading} className="w-full">
            {loading ? "Generating..." : "Generate Example Document"}
          </Button>

          <div className="rounded-md bg-muted p-3 text-xs text-muted-foreground">
            <p className="font-semibold mb-2">API Endpoint:</p>
            <code>POST /api/generate-docx</code>
          </div>
        </CardContent>
      </Card>
    </main>
  )
}
