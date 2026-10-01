'use client';

import { CalendarDays, FileSpreadsheet, ShieldCheck, UploadCloud } from 'lucide-react';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';

const uploads = [
  { name: 'Nómina', detail: 'Excel mensual con beneficiarios y movimientos' },
  { name: 'Altas', detail: 'Archivo de altas de sobrevivencia del período' },
  { name: 'Reporte revisión PBS', detail: 'Base consolidada con fechas, montos y relaciones' },
];

export default function Home() {
  return (
    <main className="min-h-screen bg-background text-foreground">
      <header className="border-b border-border/70 bg-white/80 backdrop-blur">
        <div className="mx-auto flex max-w-6xl items-center justify-between px-5 py-4 sm:px-8">
          <div className="flex items-center gap-3">
            <span className="grid size-10 place-items-center rounded-xl bg-primary text-primary-foreground shadow-sm"><FileSpreadsheet className="size-5" /></span>
            <div><p className="font-semibold tracking-tight">Revisión de pagos</p><p className="text-xs text-muted-foreground">Pensiones de sobrevivencia</p></div>
          </div>
          <Badge variant="outline" className="hidden gap-1.5 sm:flex"><ShieldCheck className="size-3.5 text-emerald-600" />Procesamiento privado</Badge>
        </div>
      </header>

      <section className="mx-auto max-w-6xl px-5 py-10 sm:px-8 sm:py-14">
        <div className="mb-9 max-w-2xl">
          <Badge className="mb-4 bg-teal-50 text-teal-800">Generador de reportes</Badge>
          <h1 className="text-3xl font-semibold tracking-[-0.035em] sm:text-4xl">Carga los archivos del período y obtén la revisión lista.</h1>
          <p className="mt-4 max-w-xl text-base leading-7 text-muted-foreground">Selecciona el análisis, adjunta los tres Excel requeridos y descarga un resultado con validaciones, estados e indexaciones.</p>
        </div>

        <div className="grid items-start gap-6 lg:grid-cols-[1.45fr_.75fr]">
          <Card className="border-0 shadow-[0_20px_55px_rgba(15,43,43,.08)] ring-1 ring-border">
            <CardHeader className="border-b">
              <CardTitle className="flex items-center gap-2 text-lg"><UploadCloud className="size-5 text-teal-700" />Archivos requeridos</CardTitle>
              <CardDescription>Se aceptan archivos .xlsx, .xlsm y .xls. Los tres deben pertenecer al mismo período.</CardDescription>
            </CardHeader>
            <CardContent className="grid gap-3 pt-1">
              {uploads.map((item, index) => (
                <label key={item.name} className="group flex cursor-pointer items-center gap-4 rounded-xl border border-dashed border-border bg-muted/25 p-4 transition hover:border-teal-600/50 hover:bg-teal-50/40">
                  <span className="grid size-9 shrink-0 place-items-center rounded-lg bg-white text-sm font-semibold text-teal-800 ring-1 ring-border">{index + 1}</span>
                  <span className="min-w-0 flex-1"><span className="block font-medium">{item.name}</span><span className="block truncate text-sm text-muted-foreground">{item.detail}</span></span>
                  <span className="text-sm font-medium text-teal-700">Seleccionar</span>
                  <input className="sr-only" type="file" accept=".xlsx,.xlsm,.xls" />
                </label>
              ))}
            </CardContent>
          </Card>

          <div className="grid gap-5">
            <Card className="border-0 ring-1 ring-border">
              <CardHeader><CardTitle>¿Qué deseas generar?</CardTitle></CardHeader>
              <CardContent className="grid gap-3">
                <label className="flex cursor-pointer gap-3 rounded-xl border border-teal-700 bg-teal-50/60 p-4"><input defaultChecked name="report" type="radio" className="mt-1 accent-teal-700" /><span><span className="block font-medium">Pendientes y no pendientes</span><span className="mt-1 block text-sm text-muted-foreground">Incluye estado, montos e indexaciones.</span></span></label>
                <label className="flex cursor-pointer gap-3 rounded-xl border p-4"><input name="report" type="radio" className="mt-1 accent-teal-700" /><span><span className="block font-medium">Retiradas</span><span className="mt-1 block text-sm text-muted-foreground">Valida edad y plazo de la pensión.</span></span></label>
              </CardContent>
            </Card>
            <Card className="border-0 bg-[#123d3a] text-white ring-0">
              <CardContent className="pt-1">
                <div className="mb-5 flex items-center gap-3 text-sm text-white/70"><CalendarDays className="size-4" />Fecha de corte: hoy</div>
                <Button disabled className="h-11 w-full bg-[#d6f36a] text-[#18322e] hover:bg-[#e2f892]">Generar reporte</Button>
                <p className="mt-3 text-center text-xs text-white/55">Adjunta los tres archivos para continuar.</p>
              </CardContent>
            </Card>
          </div>
        </div>
      </section>
    </main>
  );
}
