
"use client";

import { useState } from "react";
import { aiAssistFormulaCreation, AIAssistFormulaCreationOutput } from "@/ai/flows/ai-assist-formula-creation";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Sparkles, Loader2, Copy, Check } from "lucide-react";
import { useToast } from "@/hooks/use-toast";

export default function AssistantPage() {
  const [description, setDescription] = useState("");
  const [result, setResult] = useState<AIAssistFormulaCreationOutput | null>(null);
  const [loading, setLoading] = useState(false);
  const [copied, setCopied] = useState(false);
  const { toast } = useToast();

  const handleAsk = async () => {
    if (!description.trim()) return;
    setLoading(true);
    try {
      const out = await aiAssistFormulaCreation({ description });
      setResult(out);
    } catch (error) {
      toast({ variant: "destructive", title: "AI Error", description: "Failed to generate formula." });
    } finally {
      setLoading(false);
    }
  };

  const copyToClipboard = () => {
    if (result?.suggestedFormula) {
      navigator.clipboard.writeText(result.suggestedFormula);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
      toast({ title: "Copied to clipboard" });
    }
  };

  return (
    <div className="space-y-8 max-w-3xl mx-auto">
      <div className="flex flex-col gap-2">
        <h1 className="text-3xl font-bold font-headline flex items-center gap-3">
          <Sparkles className="h-8 w-8 text-primary" /> AI Formula Assistant
        </h1>
        <p className="text-muted-foreground">Describe what you want to calculate, and I'll generate the Excel formula for you.</p>
      </div>

      <Card className="border-none shadow-lg">
        <CardHeader>
          <CardTitle>What are you looking for?</CardTitle>
          <CardDescription>Be specific about cell ranges or operations (e.g., "calculate 20% tax for value in B2").</CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          <Textarea
            placeholder="e.g. Find the average of values from A2 to A50 if the category in B2 is 'Electronics'..."
            className="min-h-[120px] text-lg resize-none"
            value={description}
            onChange={(e) => setDescription(e.target.value)}
          />
          <Button onClick={handleAsk} disabled={loading || !description.trim()} className="w-full h-12 gap-2">
            {loading ? <Loader2 className="h-5 w-5 animate-spin" /> : <Sparkles className="h-5 w-5" />}
            Generate Formula
          </Button>
        </CardContent>
      </Card>

      {result && (
        <Card className="border-none shadow-xl bg-white animate-in slide-in-from-bottom-4 duration-300">
          <CardHeader className="pb-2">
            <CardTitle className="text-primary flex items-center gap-2">
              <Check className="h-5 w-5" /> Suggested Formula
            </CardTitle>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="relative group">
              <pre className="p-6 bg-muted rounded-xl font-mono text-xl overflow-x-auto text-foreground border border-border">
                {result.suggestedFormula}
              </pre>
              <Button 
                size="icon" 
                variant="outline" 
                className="absolute top-4 right-4 h-10 w-10 bg-white"
                onClick={copyToClipboard}
              >
                {copied ? <Check className="h-4 w-4 text-green-600" /> : <Copy className="h-4 w-4" />}
              </Button>
            </div>
            {result.explanation && (
              <div className="space-y-2 pt-2">
                <p className="font-bold text-sm uppercase text-muted-foreground tracking-wider">How it works</p>
                <p className="text-muted-foreground leading-relaxed">{result.explanation}</p>
              </div>
            )}
          </CardContent>
        </Card>
      )}
    </div>
  );
}
