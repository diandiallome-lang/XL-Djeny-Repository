"use client";

import { useState } from "react";
import { useRouter } from "next/navigation";
import { ref, uploadBytesResumable } from "firebase/storage";
import { collection, addDoc, serverTimestamp } from "firebase/firestore";
import { useFirestore, useStorage, useUser } from "@/firebase";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Progress } from "@/components/ui/progress";
import { getExcelMetadata, getTemplateFormulas, ExcelMetadata, FormulaPreview } from "@/lib/excel-utils";
import { useToast } from "@/hooks/use-toast";
import { Upload, FileSpreadsheet, Loader2, ArrowLeft, CheckCircle2, FunctionSquare } from "lucide-react";
import Link from "next/link";
import { errorEmitter } from "@/firebase/error-emitter";
import { FirestorePermissionError } from "@/firebase/errors";
import { Badge } from "@/components/ui/badge";

export default function NewTemplatePage() {
  const { user } = useUser();
  const db = useFirestore();
  const storage = useStorage();
  const router = useRouter();
  const { toast } = useToast();

  const [name, setName] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [metadata, setMetadata] = useState<ExcelMetadata | null>(null);
  const [formulas, setFormulas] = useState<FormulaPreview[]>([]);
  const [uploading, setUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (!selectedFile) return;

    try {
      const [meta, previews] = await Promise.all([
        getExcelMetadata(selectedFile),
        getTemplateFormulas(selectedFile),
      ]);
      setFile(selectedFile);
      setMetadata(meta);
      setFormulas(previews);
      if (!name) setName(selectedFile.name.replace(/\.xlsx$/i, ""));
    } catch {
      toast({
        variant: "destructive",
        title: "Fichier invalide",
        description: "Impossible de lire ce fichier. Assurez-vous qu'il s'agit d'un fichier .xlsx valide.",
      });
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!user || !file || !metadata || !db || !storage) return;

    setUploading(true);
    setUploadProgress(0);

    try {
      const modelId = crypto.randomUUID();
      const storagePath = `models/${user.uid}/${modelId}.xlsx`;
      const storageRef = ref(storage, storagePath);

      // Upload with resumable task so we can track progress
      await new Promise<void>((resolve, reject) => {
        const task = uploadBytesResumable(storageRef, file);
        task.on(
          "state_changed",
          (snapshot) => {
            const pct = Math.round(
              (snapshot.bytesTransferred / snapshot.totalBytes) * 100
            );
            setUploadProgress(pct);
          },
          reject,
          resolve
        );
      });

      const docData = {
        userId: user.uid,
        modelId,
        name,
        createdAt: serverTimestamp(),
        sheets: metadata.sheets,
        columns: metadata.columns,
        rows: metadata.rows,
        storagePath,
        formulaCount: formulas.length,
      };

      const modelsCollection = collection(db, "models");
      addDoc(modelsCollection, docData).catch(async () => {
        const permissionError = new FirestorePermissionError({
          path: modelsCollection.path,
          operation: "create",
          requestResourceData: docData,
        });
        errorEmitter.emit("permission-error", permissionError);
      });

      toast({ title: "Template enregistré avec succès" });
      router.push("/dashboard");
    } catch (error: any) {
      setUploading(false);
      setUploadProgress(0);
      toast({
        variant: "destructive",
        title: "Échec de l'upload",
        description: error.message,
      });
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
        <h1 className="text-3xl font-bold font-headline">Nouveau Template</h1>
      </div>

      <Card className="border-none shadow-lg">
        <CardHeader>
          <CardTitle>Détails du template</CardTitle>
          <CardDescription>
            Le template définit les formules (ligne 2) qui seront appliquées à chaque ligne de données.
          </CardDescription>
        </CardHeader>
        <form onSubmit={handleSubmit}>
          <CardContent className="space-y-6">
            <div className="space-y-2">
              <Label htmlFor="name">Nom du template</Label>
              <Input
                id="name"
                placeholder="ex : Rapport financier mensuel"
                required
                value={name}
                onChange={(e) => setName(e.target.value)}
              />
            </div>

            <div className="space-y-4">
              <Label>Fichier Excel (.xlsx)</Label>
              <div
                className={`relative border-2 border-dashed rounded-xl p-8 transition-colors ${
                  file
                    ? "border-primary/50 bg-primary/5"
                    : "border-muted-foreground/20 hover:border-primary/50"
                }`}
              >
                <input
                  type="file"
                  accept=".xlsx"
                  className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                  onChange={handleFileChange}
                  disabled={uploading}
                />
                <div className="flex flex-col items-center justify-center text-center">
                  {file ? (
                    <>
                      <div className="bg-primary/20 p-3 rounded-full mb-3">
                        <FileSpreadsheet className="h-8 w-8 text-primary" />
                      </div>
                      <p className="font-medium">{file.name}</p>
                      <p className="text-sm text-muted-foreground mt-1">
                        Prêt à être uploadé
                      </p>
                    </>
                  ) : (
                    <>
                      <div className="bg-muted p-3 rounded-full mb-3">
                        <Upload className="h-8 w-8 text-muted-foreground" />
                      </div>
                      <p className="font-medium text-lg">
                        Cliquer ou glisser pour sélectionner
                      </p>
                      <p className="text-sm text-muted-foreground mt-1">
                        Ligne 1 = en-têtes · Ligne 2 = formules (le modèle)
                      </p>
                    </>
                  )}
                </div>
              </div>
            </div>

            {metadata && (
              <div className="bg-muted/50 rounded-lg p-4 grid grid-cols-3 gap-4 border border-border">
                <div className="text-center">
                  <p className="text-xs text-muted-foreground uppercase font-bold mb-1">Feuilles</p>
                  <p className="text-lg font-semibold">{metadata.sheets.length}</p>
                </div>
                <div className="text-center border-x border-border">
                  <p className="text-xs text-muted-foreground uppercase font-bold mb-1">Colonnes</p>
                  <p className="text-lg font-semibold">{metadata.columns}</p>
                </div>
                <div className="text-center">
                  <p className="text-xs text-muted-foreground uppercase font-bold mb-1">Formules</p>
                  <p className="text-lg font-semibold">{formulas.length}</p>
                </div>
              </div>
            )}

            {formulas.length > 0 && (
              <div className="space-y-2">
                <div className="flex items-center gap-2">
                  <FunctionSquare className="h-4 w-4 text-primary" />
                  <p className="text-sm font-semibold text-muted-foreground uppercase tracking-wider">
                    Formules détectées
                  </p>
                </div>
                <div className="flex flex-wrap gap-2">
                  {formulas.map((f) => (
                    <Badge key={f.column} variant="secondary" className="font-mono text-xs">
                      Col {f.column} : {f.formula}
                    </Badge>
                  ))}
                </div>
              </div>
            )}

            {uploading && (
              <div className="space-y-2">
                <div className="flex justify-between text-sm text-muted-foreground">
                  <span>Upload en cours…</span>
                  <span>{uploadProgress}%</span>
                </div>
                <Progress value={uploadProgress} className="h-2" />
              </div>
            )}
          </CardContent>

          <div className="p-6 pt-0">
            <Button
              type="submit"
              className="w-full h-12 gap-2"
              disabled={uploading || !file}
            >
              {uploading ? (
                <>
                  <Loader2 className="h-4 w-4 animate-spin" /> Enregistrement…
                </>
              ) : (
                <>
                  <CheckCircle2 className="h-5 w-5" /> Créer le template
                </>
              )}
            </Button>
          </div>
        </form>
      </Card>
    </div>
  );
}
