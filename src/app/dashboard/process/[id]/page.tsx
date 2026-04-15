
"use client";

import { useState, useEffect } from "react";
import { useParams, useRouter } from "next/navigation";
import { doc, getDoc } from "firebase/firestore";
import { ref, getBytes } from "firebase/storage";
import { db, storage } from "@/lib/firebase";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { applyTemplateToData } from "@/lib/excel-utils";
import { useToast } from "@/hooks/use-toast";
import { Loader2, FileSpreadsheet, Upload, Download, CheckCircle2, ArrowLeft } from "lucide-react";
import Link from "next/link";

export default function ProcessPage() {
  const { id } = useParams() as { id: string };
  const [template, setTemplate] = useState<any>(null);
  const [dataFile, setDataFile] = useState<File | null>(null);
  const [processing, setProcessing] = useState(false);
  const [resultBlob, setResultBlob] = useState<Blob | null>(null);
  const { toast } = useToast();

  useEffect(() => {
    const fetchTemplate = async () => {
      const docRef = doc(db, "models", id);
      const docSnap = await getDoc(docRef);
      if (docSnap.exists()) {
        setTemplate(docSnap.data());
      }
    };
    fetchTemplate();
  }, [id]);

  const handleProcess = async () => {
    if (!template || !dataFile) return;

    setProcessing(true);
    try {
      // 1. Fetch template file from storage
      const templateRef = ref(storage, template.storagePath);
      const templateBuffer = await getBytes(templateRef);
      const templateFile = new File([templateBuffer], "template.xlsx");

      // 2. Apply formulas client-side
      const blob = await applyTemplateToData(templateFile, dataFile);
      setResultBlob(blob);
      toast({ title: "Processing complete!" });
    } catch (error: any) {
      toast({
        variant: "destructive",
        title: "Processing Failed",
        description: error.message,
      });
    } finally {
      setProcessing(false);
    }
  };

  const handleDownload = () => {
    if (!resultBlob) return;
    const url = URL.createObjectURL(resultBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `processed_${dataFile?.name || "data.xlsx"}`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  if (!template) {
    return (
      <div className="flex h-[60vh] items-center justify-center">
        <Loader2 className="h-8 w-8 animate-spin text-primary" />
      </div>
    );
  }

  return (
    <div className="space-y-8 max-w-2xl mx-auto">
       <div className="flex items-center gap-4">
        <Button asChild variant="ghost" size="icon">
          <Link href="/dashboard">
            <ArrowLeft className="h-5 w-5" />
          </Link>
        </Button>
        <h1 className="text-3xl font-bold font-headline">Apply Template</h1>
      </div>

      <div className="grid grid-cols-1 gap-6">
        <Card className="border-none shadow-md">
          <CardHeader>
            <div className="flex items-center gap-3 mb-2">
               <div className="p-2 bg-primary/10 text-primary rounded-lg">
                <FileSpreadsheet className="h-5 w-5" />
               </div>
               <CardTitle>{template.name}</CardTitle>
            </div>
            <CardDescription>
              This template will apply formulas from row 2 to all data rows in your new file.
            </CardDescription>
          </CardHeader>
        </Card>

        <Card className={`border-none shadow-lg transition-all ${resultBlob ? 'opacity-50 pointer-events-none' : ''}`}>
          <CardHeader>
            <CardTitle>Upload Raw Data</CardTitle>
            <CardDescription>Choose the .xlsx file you want to process.</CardDescription>
          </CardHeader>
          <CardContent>
            <div className={`relative border-2 border-dashed rounded-xl p-8 transition-colors ${dataFile ? 'border-primary/50 bg-primary/5' : 'border-muted-foreground/20 hover:border-primary/50'}`}>
              <input
                type="file"
                accept=".xlsx"
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                onChange={(e) => setDataFile(e.target.files?.[0] || null)}
              />
              <div className="flex flex-col items-center justify-center text-center">
                {dataFile ? (
                  <>
                    <div className="bg-primary/20 p-3 rounded-full mb-3">
                      <FileSpreadsheet className="h-8 w-8 text-primary" />
                    </div>
                    <p className="font-medium">{dataFile.name}</p>
                    <p className="text-sm text-muted-foreground mt-1">Ready to process</p>
                  </>
                ) : (
                  <>
                    <div className="bg-muted p-3 rounded-full mb-3">
                      <Upload className="h-8 w-8 text-muted-foreground" />
                    </div>
                    <p className="font-medium text-lg">Select Raw Data File</p>
                    <p className="text-sm text-muted-foreground mt-1">Structure must match the template columns</p>
                  </>
                )}
              </div>
            </div>
          </CardContent>
          <CardFooter>
            <Button 
              className="w-full h-12 gap-2" 
              disabled={!dataFile || processing} 
              onClick={handleProcess}
            >
              {processing ? (
                <>
                  <Loader2 className="h-4 w-4 animate-spin" /> Processing...
                </>
              ) : (
                <>
                  <Play className="h-4 w-4" /> Run Processing
                </>
              )}
            </Button>
          </CardFooter>
        </Card>

        {resultBlob && (
          <Card className="border-none shadow-xl bg-primary text-primary-foreground">
            <CardHeader>
              <div className="flex items-center gap-3">
                <CheckCircle2 className="h-6 w-6" />
                <CardTitle>Processing Successful</CardTitle>
              </div>
            </CardHeader>
            <CardContent>
              <p>Your treated Excel file is ready. Formulas have been applied to all rows.</p>
            </CardContent>
            <CardFooter className="flex gap-3">
              <Button onClick={handleDownload} variant="secondary" className="flex-1 h-12 gap-2 font-bold">
                <Download className="h-5 w-5" /> Download Result
              </Button>
              <Button onClick={() => { setResultBlob(null); setDataFile(null); }} variant="outline" className="text-primary-foreground hover:text-primary">
                Reset
              </Button>
            </CardFooter>
          </Card>
        )}
      </div>
    </div>
  );
}
