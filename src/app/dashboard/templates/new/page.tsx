
"use client";

import { useState } from "react";
import { useRouter } from "next/navigation";
import { ref, uploadBytes } from "firebase/storage";
import { collection, addDoc, serverTimestamp } from "firebase/firestore";
import { useAuth, useFirestore, useStorage, useUser } from "@/firebase";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { getExcelMetadata, ExcelMetadata } from "@/lib/excel-utils";
import { useToast } from "@/hooks/use-toast";
import { Upload, FileSpreadsheet, Loader2, ArrowLeft, CheckCircle2 } from "lucide-react";
import Link from "next/link";

export default function NewTemplatePage() {
  const { user } = useUser();
  const db = useFirestore();
  const storage = useStorage();
  const router = useRouter();
  const { toast } = useToast();
  
  const [name, setName] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [metadata, setMetadata] = useState<ExcelMetadata | null>(null);
  const [uploading, setUploading] = useState(false);

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (!selectedFile) return;

    try {
      const meta = await getExcelMetadata(selectedFile);
      setFile(selectedFile);
      setMetadata(meta);
      if (!name) setName(selectedFile.name.replace(".xlsx", ""));
    } catch (err) {
      toast({
        variant: "destructive",
        title: "Invalid File",
        description: "Could not extract metadata. Please ensure it is a valid .xlsx file.",
      });
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!user || !file || !metadata || !db || !storage) return;

    setUploading(true);
    try {
      const modelId = crypto.randomUUID();
      const storagePath = `models/${user.uid}/${modelId}.xlsx`;
      const storageRef = ref(storage, storagePath);
      
      await uploadBytes(storageRef, file);
      
      await addDoc(collection(db, "models"), {
        userId: user.uid,
        modelId,
        name,
        createdAt: serverTimestamp(),
        sheets: metadata.sheets,
        columns: metadata.columns,
        rows: metadata.rows,
        storagePath,
      });

      toast({ title: "Template saved successfully" });
      router.push("/dashboard");
    } catch (error: any) {
      toast({
        variant: "destructive",
        title: "Upload Failed",
        description: error.message,
      });
    } finally {
      setUploading(false);
    }
  };

  return (
    <div className="space-y-8 max-w-2xl mx-auto">
      <div className="flex items-center gap-4">
        <Button asChild variant="ghost" size="icon">
          <Link href="/dashboard">
            <ArrowLeft className="h-5 w-5" />
          </Link>
        </Button>
        <h1 className="text-3xl font-bold font-headline">Upload Template</h1>
      </div>

      <Card className="border-none shadow-lg">
        <CardHeader>
          <CardTitle>Template Details</CardTitle>
          <CardDescription>
            Templates define the rules and formulas that will be applied to your raw data.
          </CardDescription>
        </CardHeader>
        <form onSubmit={handleSubmit}>
          <CardContent className="space-y-6">
            <div className="space-y-2">
              <Label htmlFor="name">Template Name</Label>
              <Input
                id="name"
                placeholder="e.g. Monthly Financial Report"
                required
                value={name}
                onChange={(e) => setName(e.target.value)}
              />
            </div>

            <div className="space-y-4">
              <Label>Excel File (.xlsx)</Label>
              <div className={`relative border-2 border-dashed rounded-xl p-8 transition-colors ${file ? 'border-primary/50 bg-primary/5' : 'border-muted-foreground/20 hover:border-primary/50'}`}>
                <input
                  type="file"
                  accept=".xlsx"
                  className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                  onChange={handleFileChange}
                />
                <div className="flex flex-col items-center justify-center text-center">
                  {file ? (
                    <>
                      <div className="bg-primary/20 p-3 rounded-full mb-3">
                        <FileSpreadsheet className="h-8 w-8 text-primary" />
                      </div>
                      <p className="font-medium">{file.name}</p>
                      <p className="text-sm text-muted-foreground mt-1">Ready to upload</p>
                    </>
                  ) : (
                    <>
                      <div className="bg-muted p-3 rounded-full mb-3">
                        <Upload className="h-8 w-8 text-muted-foreground" />
                      </div>
                      <p className="font-medium text-lg">Click or drag to select</p>
                      <p className="text-sm text-muted-foreground mt-1">Upload a template with formulas in the second row</p>
                    </>
                  )}
                </div>
              </div>
            </div>

            {metadata && (
              <div className="bg-muted/50 rounded-lg p-4 grid grid-cols-3 gap-4 border border-border">
                <div className="text-center">
                  <p className="text-xs text-muted-foreground uppercase font-bold mb-1">Sheets</p>
                  <p className="text-lg font-semibold">{metadata.sheets.length}</p>
                </div>
                <div className="text-center border-x border-border">
                  <p className="text-xs text-muted-foreground uppercase font-bold mb-1">Columns</p>
                  <p className="text-lg font-semibold">{metadata.columns}</p>
                </div>
                <div className="text-center">
                  <p className="text-xs text-muted-foreground uppercase font-bold mb-1">Rows</p>
                  <p className="text-lg font-semibold">{metadata.rows}</p>
                </div>
              </div>
            )}
          </CardContent>
          <div className="p-6 pt-0">
            <Button type="submit" className="w-full h-12 gap-2" disabled={uploading || !file}>
              {uploading ? (
                <>
                  <Loader2 className="h-4 w-4 animate-spin" /> Saving Template...
                </>
              ) : (
                <>
                  <CheckCircle2 className="h-5 w-5" /> Create Template
                </>
              )}
            </Button>
          </div>
        </form>
      </Card>
    </div>
  );
}
