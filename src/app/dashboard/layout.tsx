
"use client";

import { ReactNode } from "react";
import { Sidebar, SidebarContent, SidebarFooter, SidebarHeader, SidebarMenu, SidebarMenuButton, SidebarMenuItem, SidebarProvider, SidebarRail, SidebarTrigger } from "@/components/ui/sidebar";
import { FileSpreadsheet, LayoutDashboard, Sparkles, LogOut, Plus, ShieldCheck } from "lucide-react";
import { useUser, useAuth } from "@/firebase";
import { signOut } from "firebase/auth";
import { useRouter, usePathname } from "next/navigation";
import Link from "next/link";

export default function DashboardLayout({ children }: { children: ReactNode }) {
  const { user, isAdmin, loading } = useUser();
  const auth = useAuth();
  const router = useRouter();
  const pathname = usePathname();

  const handleLogout = async () => {
    if (!auth) return;
    await signOut(auth);
    router.push("/auth");
  };

  if (loading) return null;
  if (!user) {
    router.push("/auth");
    return null;
  }

  const menuItems = [
    { icon: LayoutDashboard, label: "Templates", href: "/dashboard" },
    { icon: Plus, label: "New Template", href: "/dashboard/templates/new" },
    { icon: Sparkles, label: "AI Assistant", href: "/dashboard/assistant" },
  ];

  return (
    <SidebarProvider>
      <div className="flex min-h-screen w-full">
        <Sidebar className="border-none shadow-xl shadow-primary/10">
          <SidebarHeader className="p-4">
            <div className="flex items-center gap-3 px-2">
              <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-primary text-primary-foreground shadow-lg shadow-primary/20">
                <FileSpreadsheet className="h-5 w-5" />
              </div>
              <div className="flex flex-col">
                <span className="font-headline text-lg font-bold tracking-tight">Formulytics</span>
                {isAdmin && (
                  <div className="flex items-center gap-1 text-[10px] text-primary-foreground/70 bg-white/10 px-1.5 rounded-full w-fit mt-0.5">
                    <ShieldCheck className="h-2 w-2" />
                    <span>ADMIN</span>
                  </div>
                )}
              </div>
            </div>
          </SidebarHeader>
          <SidebarContent>
            <SidebarMenu className="px-2 py-4">
              {menuItems.map((item) => (
                <SidebarMenuItem key={item.href}>
                  <SidebarMenuButton 
                    asChild 
                    isActive={pathname === item.href}
                    className="h-10 px-4 hover:bg-sidebar-accent/50 data-[active=true]:bg-sidebar-accent"
                  >
                    <Link href={item.href}>
                      <item.icon className="mr-3 h-5 w-5" />
                      <span className="font-medium">{item.label}</span>
                    </Link>
                  </SidebarMenuButton>
                </SidebarMenuItem>
              ))}
            </SidebarMenu>
          </SidebarContent>
          <SidebarFooter className="p-4 border-t border-sidebar-border">
            <div className="mb-4 px-2">
              <p className="text-xs text-sidebar-foreground/50 truncate">{user.email}</p>
            </div>
            <SidebarMenuButton onClick={handleLogout} className="text-destructive hover:bg-destructive/10 hover:text-destructive">
              <LogOut className="mr-3 h-5 w-5" />
              <span className="font-medium">Sign Out</span>
            </SidebarMenuButton>
          </SidebarFooter>
          <SidebarRail />
        </Sidebar>
        <main className="flex-1 overflow-auto bg-background p-6 md:p-8">
          <div className="mx-auto max-w-6xl">
            <SidebarTrigger className="md:hidden mb-6" />
            {children}
          </div>
        </main>
      </div>
    </SidebarProvider>
  );
}
