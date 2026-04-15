
"use client";

import { useMemo } from "react";
import { collection, query, where, deleteDoc, doc } from "firebase/firestore";
import { useUser, useFirestore, useCollection } from "@/firebase";
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Plus, Download, Trash2, Play, FileSpreadsheet, Loader2 } from "lucide-react";
import Link from "next/link";
import { format } from "date-fns";
import { getStorage, ref, getDownloadURL } from "firebase/storage";
import { useToast } from "@/hooks/use-toast";

interface Template {
  id: string;
  name: string;
  sheets: string[];
  columns: number;
  rows: number;
  createdAt: any;
  storagePath: string;
}

export default function TemplatesPage() {
  const { user } = useUser();
  const db = useFirestore();
  const { toast } = useToast();

  const templatesQuery = useMemo(() => {
    if (!db || !user) return null;
    return query(collection(db, "models"), where("userId", "==", user.uid));
  }, [db, user]);

  const { data: templates, loading } = useCollection<Template>(templatesQuery);

  const handleDelete = async (id: string) => {
    if (!db) return;
    try {
      await deleteDoc(doc(db, "models", id));
      toast({ title: "Template deleted successfully" });
    } catch (error) {
      toast({ variant: "destructive", title: "Error deleting template" });
    }
  };

  const handleDownload = async (template: Template) => {
    if (!db) return;
    const storage = getStorage();
    try {
      const url = await getDownloadURL(ref(storage, template.storagePath));
      window.open(url, "_blank");
    } catch (error) {
      toast({ variant: "destructive", title: "Error downloading file" });
    }
  };

  if (loading) {
    return (
      <div className="flex h-[60vh] items-center justify-center">
        <Loader2 className="h-8 w-8 animate-spin text-primary" />
      </div>
    );
  }

  const sortedTemplates = [...(templates || [])].sort((a, b) => 
    (b.createdAt?.seconds || 0) - (a.createdAt?.seconds || 0)
  );

  return (
    <div className="space-y-8">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-3xl font-bold tracking-tight font-headline">My Templates</h1>
          <p className="text-muted-foreground">Manage and apply your Excel processing rules.</p>
        </div>
        <Button asChild className="gap-2 shadow-lg shadow-primary/20">
          <Link href="/dashboard/templates/new">
            <Plus className="h-4 w-4" /> New Template
          </Link>
        </Button>
      </div>

      {sortedTemplates.length === 0 ? (
        <Card className="flex flex-col items-center justify-center p-12 text-center border-dashed border-2">
          <div className="rounded-full bg-muted p-4 mb-4">
            <FileSpreadsheet className="h-8 w-8 text-muted-foreground" />
          </div>
          <CardTitle>No templates yet</CardTitle>
          <CardDescription className="mt-2 max-w-sm">
            Upload your first Excel template to start processing raw data automatically.
          </CardDescription>
          <Button asChild variant="outline" className="mt-6">
            <Link href="/dashboard/templates/new">Upload Template</Link>
          </Button>
        </Card>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {sortedTemplates.map((t) => (
            <Card key={t.id} className="group hover:shadow-xl transition-all duration-300 border-none shadow-md overflow-hidden bg-white">
              <CardHeader className="pb-4">
                <div className="flex items-start justify-between">
                  <div className="p-2 bg-primary/5 rounded-lg text-primary mb-2">
                    <FileSpreadsheet className="h-6 w-6" />
                  </div>
                  <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                    <Button variant="ghost" size="icon" className="h-8 w-8 text-muted-foreground hover:text-primary" onClick={() => handleDownload(t)}>
                      <Download className="h-4 w-4" />
                    </Button>
                    <Button variant="ghost" size="icon" className="h-8 w-8 text-muted-foreground hover:text-destructive" onClick={() => handleDelete(t.id)}>
                      <Trash2 className="h-4 w-4" />
                    </Button>
                  </div>
                </div>
                <CardTitle className="text-xl font-headline truncate">{t.name}</CardTitle>
                <CardDescription>
                  Added {t.createdAt?.seconds ? format(new Date(t.createdAt.seconds * 1000), "MMM d, yyyy") : "Recently"}
                </CardDescription>
              </CardHeader>
              <CardContent className="pb-4">
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <p className="text-xs text-muted-foreground uppercase font-semibold">Sheets</p>
                    <p className="font-medium">{t.sheets?.length || 0}</p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-xs text-muted-foreground uppercase font-semibold">Columns</p>
                    <p className="font-medium">{t.columns || 0}</p>
                  </div>
                </div>
              </CardContent>
              <CardFooter className="pt-0">
                <Button asChild className="w-full gap-2 group/btn" variant="secondary">
                  <Link href={`/dashboard/process/${t.id}`}>
                    <Play className="h-4 w-4 transition-transform group-hover/btn:scale-110" /> Apply Template
                  </Link>
                </Button>
              </CardFooter>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}
