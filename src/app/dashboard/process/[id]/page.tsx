
"use client";

import { useState, useEffect } from "react";
import { useParams } from "next/navigation";
import { doc, getDoc } from "firebase/firestore";
import { ref, getBytes } from "firebase/storage";
import { useFirestore, useStorage } from "@/firebase";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardFooter,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { applyTemplateToData, getTemplateFormulas, FormulaPreview } from "@/lib/excel-utils";
import { useToast } from "@/hooks/use-toast";
import {
  Loader2,
  FileSpreadsheet,
  Upload,
  Download,
  CheckCircle2,
  ArrowLeft,
  Play,
  FunctionSquare,
} from "lucide-react";
import Link from "next/link";

export default function ProcessPage() {
  const { id } = useParams() as { id: string };
  const db = useFirestore();
  const storage = useStorage();

  const [template, setTemplate] = useState<any>(null);
  const [templateFormulas, setTemplateFormulas] = useState<FormulaPreview[]>([]);
  const [dataFile, setDataFile] = useState<File | null>(null);
  const [processing, setProcessing] = useState(false);
  const [resultBlob, setResultBlob] = useState<Blob | null>(null);
  const { toast } = useToast();

  useEffect(() => {
    const fetchTemplate = async () => {
      if (!db || !storage) return;
      const docRef = doc(db, "models", id);
      const docSnap = await getDoc(docRef);
      if (!docSnap.exists()) return;

      const data = docSnap.data();
      setTemplate(data);

      // Load formulas preview from the stored file
      try {
        const templateRef = ref(storage, data.storagePath);
        const templateBuffer = await getBytes(templateRef);
        const templateFile = new File([templateBuffer], "template.xlsx");
        const previews = await getTemplateFormulas(templateFile);
        setTemplateFormulas(previews);
      } catch {
        // Non-blocking — formula preview is informational only
      }
    };
    fetchTemplate();
  }, [id, db, storage]);

  const handleProcess = async () => {
    if (!template || !dataFile || !storage) return;

    setProcessing(true);
    try {
      const templateRef = ref(storage, template.storagePath);
      const templateBuffer = await getBytes(templateRef);
      const templateFile = new File([templateBuffer], "template.xlsx");

      const blob = await applyTemplateToData(templateFile, dataFile);
      setResultBlob(blob);
      toast({ title: "Traitement terminé !" });
    } catch (error: any) {
      toast({
        variant: "destructive",
        title: "Échec du traitement",
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
    a.download = `traite_${dataFile?.name || "data.xlsx"}`;
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
        <h1 className="text-3xl font-bold font-headline">Appliquer le template</h1>
      </div>

      <div className="grid grid-cols-1 gap-6">
        {/* Template info card */}
        <Card className="border-none shadow-md">
          <CardHeader>
            <div className="flex items-center gap-3 mb-1">
              <div className="p-2 bg-primary/10 text-primary rounded-lg">
                <FileSpreadsheet className="h-5 w-5" />
              </div>
              <CardTitle>{template.name}</CardTitle>
            </div>
            <CardDescription>
              Les formules de la ligne 2 du template seront reproduites sur chaque ligne de
              données de votre fichier.
            </CardDescription>
          </CardHeader>

          {templateFormulas.length > 0 && (
            <CardContent className="pt-0">
              <div className="space-y-2">
                <div className="flex items-center gap-2">
                  <FunctionSquare className="h-4 w-4 text-primary" />
                  <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wider">
                    Formules du template
                  </p>
                </div>
                <div className="flex flex-wrap gap-2">
                  {templateFormulas.map((f) => (
                    <Badge key={f.column} variant="secondary" className="font-mono text-xs">
                      Col {f.column} : {f.formula}
                    </Badge>
                  ))}
                </div>
              </div>
            </CardContent>
          )}
        </Card>

        {/* Upload raw data */}
        <Card
          className={`border-none shadow-lg transition-all ${
            resultBlob ? "opacity-50 pointer-events-none" : ""
          }`}
        >
          <CardHeader>
            <CardTitle>Charger les données brutes</CardTitle>
            <CardDescription>
              Sélectionnez le fichier .xlsx contenant vos données. La structure des colonnes
              doit correspondre à celle du template.
            </CardDescription>
          </CardHeader>
          <CardContent>
            <div
              className={`relative border-2 border-dashed rounded-xl p-8 transition-colors ${
                dataFile
                  ? "border-primary/50 bg-primary/5"
                  : "border-muted-foreground/20 hover:border-primary/50"
              }`}
            >
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
                    <p className="text-sm text-muted-foreground mt-1">Prêt à traiter</p>
                  </>
                ) : (
                  <>
                    <div className="bg-muted p-3 rounded-full mb-3">
                      <Upload className="h-8 w-8 text-muted-foreground" />
                    </div>
                    <p className="font-medium text-lg">Sélectionner le fichier de données</p>
                    <p className="text-sm text-muted-foreground mt-1">
                      Ligne 1 = en-têtes · Lignes 2+ = données brutes
                    </p>
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
                  <Loader2 className="h-4 w-4 animate-spin" /> Traitement en cours…
                </>
              ) : (
                <>
                  <Play className="h-4 w-4" /> Lancer le traitement
                </>
              )}
            </Button>
          </CardFooter>
        </Card>

        {/* Result card */}
        {resultBlob && (
          <Card className="border-none shadow-xl bg-primary text-primary-foreground">
            <CardHeader>
              <div className="flex items-center gap-3">
                <CheckCircle2 className="h-6 w-6" />
                <CardTitle>Traitement réussi</CardTitle>
              </div>
            </CardHeader>
            <CardContent>
              <p>
                Votre fichier Excel traité est prêt. Les formules du template ont été
                appliquées et ajustées pour chaque ligne de données.
              </p>
            </CardContent>
            <CardFooter className="flex gap-3">
              <Button
                onClick={handleDownload}
                variant="secondary"
                className="flex-1 h-12 gap-2 font-bold"
              >
                <Download className="h-5 w-5" /> Télécharger le résultat
              </Button>
              <Button
                onClick={() => {
                  setResultBlob(null);
                  setDataFile(null);
                }}
                variant="outline"
                className="text-primary-foreground hover:text-primary"
              >
                Réinitialiser
              </Button>
            </CardFooter>
          </Card>
        )}
      </div>
    </div>
  );
}
