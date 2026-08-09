import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  AlertTriangle,
  Plus,
  Search,
  X,
  Edit2,
  Trash2,
  Building2,
  GraduationCap,
  MapPin,
  FileText,
  Clock,
  ChevronDown,
  ChevronRight,
  Landmark,
  UserMinus,
  RefreshCw,
  PauseCircle,
  CheckCircle2,
  Download,
  Printer,
  Save,
  UploadCloud,
  Eye,
  EyeOff,
} from "lucide-react";

// ---------- Utilidades de fecha ----------
const toDate = (s) => (s ? new Date(s + "T00:00:00") : null);
const fmt = (s) => {
  if (!s) return "—";
  const d = toDate(s);
  return d.toLocaleDateString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric" });
};
const addDays = (date, days) => {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
};
const diffDaysInclusive = (a, b) => Math.round((toDate(b) - toDate(a)) / 86400000) + 1;
const daysBetween = (a, b) => Math.round((b - a) / 86400000);
const monthLabel = (y, m) =>
  new Date(y, m, 1).toLocaleDateString("es-AR", { month: "long", year: "numeric" });
const todayStr = () => new Date().toISOString().slice(0, 10);

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);

// ---------- Lógica de estado ----------
function fechaFinBase(p) {
  return p.renovacion?.fin || p.periodo1.fin;
}
function diasSuspendidos(p) {
  return (p.suspensiones || []).reduce((acc, s) => acc + (s.inicio && s.fin ? diffDaysInclusive(s.inicio, s.fin) : 0), 0);
}
function fechaVencimientoEfectiva(p) {
  const base = toDate(fechaFinBase(p));
  return addDays(base, diasSuspendidos(p));
}
function suspensionActiva(p, ref) {
  return (p.suspensiones || []).find((s) => {
    if (!s.inicio || !s.fin) return false;
    return toDate(s.inicio) <= ref && ref <= toDate(s.fin);
  });
}
function getEstado(p, ref = new Date()) {
  ref.setHours(0, 0, 0, 0);
  if (p.renuncia?.fecha && toDate(p.renuncia.fecha) <= ref) return "renuncia";
  const susp = suspensionActiva(p, ref);
  if (susp) return "suspendida";
  const efectiva = fechaVencimientoEfectiva(p);
  const restantes = daysBetween(ref, efectiva);
  if (p.renuncia?.fecha && toDate(p.renuncia.fecha) > ref) {
    // renuncia futura ya cargada: sigue vigente hasta esa fecha
  }
  if (restantes < 0) return "vencida";
  if (restantes <= 45) return p.renovacion ? "por_vencer_final" : "por_vencer_renovacion";
  return "vigente";
}
const ESTADOS = {
  vigente: { label: "Vigente", color: "#2F6E5E", bg: "#EAF3EF" },
  por_vencer_renovacion: { label: "Por vencer · iniciar renovación", color: "#B8862F", bg: "#FBF3E4" },
  por_vencer_final: { label: "Por vencer · finalización", color: "#B8862F", bg: "#FBF3E4" },
  suspendida: { label: "Suspendida", color: "#6B5B95", bg: "#EFEDF5" },
  vencida: { label: "Vencida", color: "#A6432D", bg: "#F6E9E6" },
  renuncia: { label: "Renuncia", color: "#8A4B3B", bg: "#F3E9E5" },
};

// ---------- Novedades derivadas ----------
function novedadesDe(p) {
  const eventos = [];
  if (p.periodo1?.inicio) {
    eventos.push({
      fecha: p.periodo1.inicio,
      tipo: "alta",
      detalle: `Alta · 1er período${p.expediente ? ` (Exp. ${p.expediente})` : ""}`,
    });
  }
  if (p.renovacion?.inicio) {
    eventos.push({ fecha: p.renovacion.inicio, tipo: "renovacion", detalle: "Renovación de convenio" });
  }
  (p.suspensiones || []).forEach((s) => {
    if (s.inicio)
      eventos.push({
        fecha: s.inicio,
        tipo: "suspension_inicio",
        detalle: `Inicio de suspensión${s.motivo ? ` · ${s.motivo}` : ""}`,
      });
    if (s.fin)
      eventos.push({ fecha: s.fin, tipo: "suspension_fin", detalle: "Fin de suspensión" });
  });
  if (p.renuncia?.fecha) {
    eventos.push({
      fecha: p.renuncia.fecha,
      tipo: "renuncia",
      detalle: `Baja · Renuncia${p.renuncia.nota ? ` (${p.renuncia.nota})` : ""}`,
    });
  } else {
    const efectiva = fechaVencimientoEfectiva(p);
    const ref = new Date();
    ref.setHours(0, 0, 0, 0);
    if (efectiva < ref) {
      const iso = efectiva.toISOString().slice(0, 10);
      eventos.push({
        fecha: iso,
        tipo: "baja_vencimiento",
        detalle: p.renovacion ? "Baja · Fin de contrato (finalización de renovación)" : "Baja · Fin de contrato (1er período, sin renovación)",
      });
    }
  }
  return eventos;
}

const TIPO_LABEL = {
  alta: { label: "Alta", icon: Plus, color: "#2F6E5E" },
  renovacion: { label: "Renovación", icon: RefreshCw, color: "#2F6E7A" },
  suspension_inicio: { label: "Inicio suspensión", icon: PauseCircle, color: "#6B5B95" },
  suspension_fin: { label: "Fin suspensión", icon: CheckCircle2, color: "#6B5B95" },
  renuncia: { label: "Baja (renuncia)", icon: UserMinus, color: "#A6432D" },
  baja_vencimiento: { label: "Baja (vencimiento)", icon: Clock, color: "#A6432D" },
};

// ---------- Exportación a Excel ----------
function filaExport(p) {
  return {
    DNI: p.dni || "",
    Nombre: p.nombre || "",
    "Lugar de trabajo": p.lugarTrabajo || "",
    "Habilitación pagadora": p.habilitacionPagadora || "",
    Jurisdicción: p.jurisdiccion || "",
    "Universidad/Instituto": p.universidad || "",
    Carrera: p.carrera || "",
    Expediente: p.expediente || "",
    "1er período - desde": fmt(p.periodo1?.inicio),
    "1er período - hasta": fmt(p.periodo1?.fin),
    "1er período - convenio": p.periodo1?.convenio || "",
    "Renovación - desde": p.renovacion?.inicio ? fmt(p.renovacion.inicio) : "",
    "Renovación - hasta": p.renovacion?.fin ? fmt(p.renovacion.fin) : "",
    "Renovación - convenio": p.renovacion?.convenio || "",
    "Renuncia - fecha": p.renuncia?.fecha ? fmt(p.renuncia.fecha) : "",
    "Renuncia - nota": p.renuncia?.nota || "",
    "Renuncia - observaciones": p.renuncia?.observaciones || "",
    Suspensiones: (p.suspensiones || [])
      .map((s) => `${fmt(s.inicio)} a ${fmt(s.fin)}${s.motivo ? " (" + s.motivo + ")" : ""}${s.convenio ? " · convenio " + s.convenio : ""}`)
      .join(" | "),
    Estado: ESTADOS[p._estado || getEstado(p)]?.label || "",
    "Vencimiento efectivo": fmt((p._vencEfectiva || fechaVencimientoEfectiva(p)).toISOString().slice(0, 10)),
  };
}

function exportExcel({ detalle, grupos, nombreArchivo, hojaDetalle = "Detalle" }) {
  const wb = XLSX.utils.book_new();
  if (grupos && grupos.length) {
    const filasResumen = grupos.map((g) => ({ Criterio: g.key, Cantidad: g.count }));
    const wsResumen = XLSX.utils.json_to_sheet(filasResumen);
    XLSX.utils.book_append_sheet(wb, wsResumen, "Resumen");
  }
  const wsDetalle = XLSX.utils.json_to_sheet(detalle.map(filaExport));
  XLSX.utils.book_append_sheet(wb, wsDetalle, hojaDetalle);
  XLSX.writeFile(wb, nombreArchivo);
}

function filaLiquidacion(f, eeMap) {
  return {
    DNI: f.p.dni,
    Nombre: f.p.nombre,
    "Habilitación pagadora": f.p.habilitacionPagadora,
    Jurisdicción: f.p.jurisdiccion,
    "Universidad/Instituto": f.p.universidad,
    "Lugar de trabajo": f.p.lugarTrabajo,
    Período: f.periodoTipo === "renovacion" ? "Renovación" : "Período inicial",
    "Vigencia desde": fmt(f.vigenciaDesde),
    "Vigencia hasta": fmt(f.vigenciaHasta),
    "Días trabajados": f.dias,
    "Mes completo": f.esMesCompleto ? "Sí" : "No",
    "Asignación estímulo": Number(f.monto.toFixed(2)),
    [`IAPOS (${IAPOS_PCT}%)`]: Number(f.iapos.toFixed(2)),
    "% Gasto admin.": f.gastoPct,
    "Gasto administrativo": Number(f.gastoAdmin.toFixed(2)),
    Total: Number(totalACargoHabilitacion(f).toFixed(2)),
    "N° EE": (eeMap && eeMap[f.p.dni]) || "",
    Observación: [f.centralizada ? "Gasto administrativo abonado por la habilitación centralizada" : "", f.motivoDiff || ""].filter(Boolean).join(" · "),
  };
}

function exportLiquidacionExcel(liq, nombreArchivo, etiqueta, eeMap) {
  const wb = XLSX.utils.book_new();

  const filasSubtotales = liq.porHabilitacion.map((g) => ({
    "Habilitación pagadora": g.habilitacion,
    "Subtotal (asignación + IAPOS)": Number(g.subtotalAsigIapos.toFixed(2)),
    "Gasto administrativo": Number(g.gastoAdminTotal.toFixed(2)),
    "Total a pagar": Number(g.total.toFixed(2)),
  }));
  const wsSubtotales = XLSX.utils.json_to_sheet(filasSubtotales);
  XLSX.utils.book_append_sheet(wb, wsSubtotales, "Subtotales por habilitación");

  const wsDetalle = XLSX.utils.json_to_sheet(liq.filas.map((f) => filaLiquidacion(f, eeMap)));
  XLSX.utils.book_append_sheet(wb, wsDetalle, "Detalle " + etiqueta);

  if (liq.centralizada.length) {
    const filasCentral = [];
    liq.centralizada.forEach((g) => {
      g.filas.forEach((f) => {
        filasCentral.push({
          Jurisdicción: g.jurisdiccion,
          DNI: f.p.dni,
          Nombre: f.p.nombre,
          "Habilitación pagadora (original)": f.p.habilitacionPagadora,
          "Lugar de trabajo": f.p.lugarTrabajo,
          "Universidad/Instituto": f.p.universidad,
          "% Gasto admin.": f.gastoPct,
          "Gasto administrativo": Number(f.gastoAdmin.toFixed(2)),
        });
      });
    });
    const wsCentral = XLSX.utils.json_to_sheet(filasCentral);
    XLSX.utils.book_append_sheet(wb, wsCentral, "Centralizada (gasto admin.)");
  }

  XLSX.writeFile(wb, nombreArchivo);
}

// ---------- Liquidación (Parte 2) ----------
// El catálogo de asignaciones es histórico y aditivo: puede haber más de un valor
// cargado para el mismo mes (norma retroactiva). El valor "vigente" es el más nuevo
// según fecha de carga, pero los valores anteriores no se borran.
function getAsignacion(catalogos, year, month) {
  const items = (catalogos.asignaciones || []).filter((a) => a.year === year && a.month === month);
  if (items.length === 0) return null;
  const latest = items.slice().sort((a, b) => (b.creadoEl || "").localeCompare(a.creadoEl || ""))[0];
  return latest.monto;
}

function historialAsignacion(catalogos, year, month) {
  return (catalogos.asignaciones || [])
    .filter((a) => a.year === year && a.month === month)
    .slice()
    .sort((a, b) => (b.creadoEl || "").localeCompare(a.creadoEl || ""));
}

function diasEnMes(year, month) {
  return new Date(year, month + 1, 0).getDate();
}

function calcularFilaPasantia(p, year, month, valorAsignacion, catalogos) {
  const monthStart = new Date(year, month, 1);
  const monthEnd = new Date(year, month + 1, 0);
  const inicio = toDate(p.periodo1.inicio);
  const finReal = p.renuncia?.fecha ? toDate(p.renuncia.fecha) : fechaVencimientoEfectiva(p);
  if (!inicio || !finReal) return null;
  const rangeStart = inicio > monthStart ? inicio : monthStart;
  const rangeEnd = finReal < monthEnd ? finReal : monthEnd;
  if (rangeStart > rangeEnd) return null;
  const totalDiasRango = daysBetween(rangeStart, rangeEnd) + 1;
  let suspendidos = 0;
  (p.suspensiones || []).forEach((s) => {
    if (!s.inicio || !s.fin) return;
    const sIni = toDate(s.inicio);
    const sFin = toDate(s.fin);
    const ovStart = sIni > rangeStart ? sIni : rangeStart;
    const ovEnd = sFin < rangeEnd ? sFin : rangeEnd;
    if (ovStart <= ovEnd) suspendidos += daysBetween(ovStart, ovEnd) + 1;
  });
  const diasTrabajados = Math.max(totalDiasRango - suspendidos, 0);
  if (diasTrabajados === 0) return null;
  const esMesCompleto = rangeStart.getTime() === monthStart.getTime() && rangeEnd.getTime() === monthEnd.getTime() && suspendidos === 0;
  const monto = esMesCompleto ? valorAsignacion : (valorAsignacion / 30) * diasTrabajados;
  const iapos = monto * (IAPOS_PCT / 100);
  const uni = (catalogos.universidades || []).find((u) => u.nombre === p.universidad);
  const gastoPct = uni ? Number(uni.gastoAdmin) || 0 : 0;
  const gastoAdmin = monto * (gastoPct / 100);
  const total = monto + iapos + gastoAdmin;
  const centralizada = !!uni?.centralizada;

  // Dato informativo: a qué período corresponde este mes (inicial o renovación) y su vigencia.
  let periodoTipo = "inicial";
  let vigenciaDesde = p.periodo1.inicio;
  let vigenciaHasta = p.periodo1.fin;
  if (p.renovacion?.inicio && rangeEnd >= toDate(p.renovacion.inicio)) {
    periodoTipo = "renovacion";
    vigenciaDesde = p.renovacion.inicio;
    vigenciaHasta = p.renovacion.fin;
  }

  return { p, dias: diasTrabajados, diasMes: diasEnMes(year, month), esMesCompleto, monto, iapos, gastoPct, gastoAdmin, total, centralizada, periodoTipo, vigenciaDesde, vigenciaHasta };
}

// Agrupa un conjunto de filas ya calculadas (sea una liquidación normal o un set de diferencias)
// aplicando siempre la misma regla de pago: asignación+IAPOS los paga la habilitación pagadora;
// el gasto administrativo también, salvo universidades centralizadas (lo paga la centralizada).
function agruparFilas(filas) {
  // Orden dentro de cada grupo: por universidad y, dentro de la misma universidad, alfabético por nombre.
  const ordenarFilas = (arr) =>
    arr.slice().sort((a, b) => (a.p.universidad || "").localeCompare(b.p.universidad || "") || (a.p.nombre || "").localeCompare(b.p.nombre || ""));

  const porHabMap = new Map();
  filas.forEach((f) => {
    const key = f.p.habilitacionPagadora || "(sin habilitación pagadora)";
    if (!porHabMap.has(key)) porHabMap.set(key, { habilitacion: key, filas: [], subtotalAsigIapos: 0, gastoAdminTotal: 0 });
    const g = porHabMap.get(key);
    g.filas.push(f);
    g.subtotalAsigIapos += f.monto + f.iapos;
    if (!f.centralizada) g.gastoAdminTotal += f.gastoAdmin;
  });
  const porHabilitacion = Array.from(porHabMap.values())
    .map((g) => ({ ...g, filas: ordenarFilas(g.filas), total: g.subtotalAsigIapos + g.gastoAdminTotal }))
    .sort((a, b) => a.habilitacion.localeCompare(b.habilitacion));

  const filasCentralizada = filas.filter((f) => f.centralizada);
  const centralMap = new Map();
  filasCentralizada.forEach((f) => {
    const key = f.p.jurisdiccion || "(sin jurisdicción)";
    if (!centralMap.has(key)) centralMap.set(key, { jurisdiccion: key, filas: [], subtotal: 0 });
    const g = centralMap.get(key);
    g.filas.push(f);
    g.subtotal += f.gastoAdmin;
  });
  const centralizada = Array.from(centralMap.values())
    .map((g) => ({ ...g, filas: ordenarFilas(g.filas) }))
    .sort((a, b) => a.jurisdiccion.localeCompare(b.jurisdiccion));
  const centralizadaTotal = filasCentralizada.reduce((acc, f) => acc + f.gastoAdmin, 0);

  return { filas, porHabilitacion, centralizada, centralizadaTotal };
}

function calcularLiquidacion(pasantias, catalogos, year, month, valorAsignacion) {
  const filas = pasantias.map((p) => calcularFilaPasantia(p, year, month, valorAsignacion, catalogos)).filter(Boolean);
  return agruparFilas(filas);
}

// Reconstruye, sumando todo lo confirmado hasta ahora para un mes (liquidación mensual +
// las complementarias que se hayan confirmado después), cuánto quedó pagado por pasantía.
// Esto es la "base" contra la que se compara para detectar diferencias o armar una nueva complementaria.
function liquidacionConfirmadaAcumulada(liquidaciones, year, month) {
  const registros = liquidaciones
    .filter((r) => r.year === year && r.month === month)
    .slice()
    .sort((a, b) => a.fechaConfirmacion.localeCompare(b.fechaConfirmacion));
  const acumulado = new Map();
  registros.forEach((r) => {
    r.liq.filas.forEach((f) => {
      const dni = f.p.dni;
      const prev = acumulado.get(dni) || { monto: 0, iapos: 0, gastoAdmin: 0 };
      acumulado.set(dni, {
        p: f.p,
        monto: prev.monto + f.monto,
        iapos: prev.iapos + f.iapos,
        gastoAdmin: prev.gastoAdmin + f.gastoAdmin,
        gastoPct: f.gastoPct,
        centralizada: f.centralizada,
        dias: f.dias,
        esMesCompleto: f.esMesCompleto,
        periodoTipo: f.periodoTipo,
        vigenciaDesde: f.vigenciaDesde,
        vigenciaHasta: f.vigenciaHasta,
      });
    });
  });
  const filas = Array.from(acumulado.values()).map((x) => ({ ...x, total: x.monto + x.iapos + x.gastoAdmin }));
  return { filas, tieneRegistros: registros.length > 0, ultimaFecha: registros[registros.length - 1]?.fechaConfirmacion };
}

// Compara la liquidación actual (recalculada con los datos de hoy) contra una base
// (lo ya confirmado) y arma las filas de diferencia: altas nuevas, bajas, y ajustes de monto.
function calcularDiferenciaLiquidacion(liqActual, baseFilas) {
  const mapaActual = new Map(liqActual.filas.map((f) => [f.p.dni, f]));
  const mapaBase = new Map(baseFilas.map((f) => [f.p.dni, f]));
  const dnis = new Set([...mapaActual.keys(), ...mapaBase.keys()]);
  const filasDelta = [];
  dnis.forEach((dni) => {
    const act = mapaActual.get(dni);
    const base = mapaBase.get(dni);
    const monto = (act?.monto || 0) - (base?.monto || 0);
    const iapos = (act?.iapos || 0) - (base?.iapos || 0);
    const gastoAdmin = (act?.gastoAdmin || 0) - (base?.gastoAdmin || 0);
    if (Math.abs(monto) < 0.01 && Math.abs(iapos) < 0.01 && Math.abs(gastoAdmin) < 0.01) return;
    const ref = act || base;
    filasDelta.push({
      p: ref.p,
      dias: ref.dias,
      esMesCompleto: ref.esMesCompleto,
      periodoTipo: ref.periodoTipo,
      vigenciaDesde: ref.vigenciaDesde,
      vigenciaHasta: ref.vigenciaHasta,
      gastoPct: ref.gastoPct,
      centralizada: ref.centralizada,
      monto,
      iapos,
      gastoAdmin,
      total: monto + iapos + gastoAdmin,
      motivoDiff: !base ? "alta nueva" : !act ? "ya no vigente" : "ajuste",
    });
  });
  return filasDelta;
}

function moneyFmt(n) {
  return (n || 0).toLocaleString("es-AR", { style: "currency", currency: "ARS", minimumFractionDigits: 2 });
}

// Le asigna un color consistente a cada valor de "estado" de trámite (son categorías
// libres que carga cada usuario, así que no hay un significado fijo: solo buscamos
// que cada una se distinga visualmente de las demás en el listado).
const PALETA_ESTADOS = [
  { bg: "#EAF3EF", color: "#2F6E5E" },
  { bg: "#FBF3E4", color: "#B8862F" },
  { bg: "#EFEDF5", color: "#6B5B95" },
  { bg: "#E6EEF6", color: "#2F5E8A" },
  { bg: "#F6E9E6", color: "#A6432D" },
  { bg: "#F3E9E5", color: "#8A4B3B" },
  { bg: "#EAF0E4", color: "#5A7A2E" },
];
function colorPorEstado(estado) {
  if (!estado) return PALETA_ESTADOS[0];
  let hash = 0;
  for (let i = 0; i < estado.length; i++) hash = (hash * 31 + estado.charCodeAt(i)) >>> 0;
  return PALETA_ESTADOS[hash % PALETA_ESTADOS.length];
}

// Lo que efectivamente paga la habilitación pagadora por esta fila: si la universidad
// es centralizada, el gasto administrativo lo abona la centralizada, así que no debe
// sumarse acá (aunque f.total, usado para Excel/otros cálculos, sí lo incluye).
function totalACargoHabilitacion(f) {
  return f.centralizada ? f.monto + f.iapos : f.total;
}

// ---------- Vigencia en período ----------
function vigenteEnMes(p, year, month) {
  const inicioMes = new Date(year, month, 1);
  const finMes = new Date(year, month + 1, 0);
  const inicio = toDate(p.periodo1.inicio);
  if (!inicio || inicio > finMes) return false;
  const finReal = p.renuncia?.fecha ? toDate(p.renuncia.fecha) : fechaVencimientoEfectiva(p);
  if (finReal < inicioMes) return false;
  return true;
}

// ---------- Timeline visual ----------
function Timeline({ p }) {
  const inicio = toDate(p.periodo1.inicio);
  const finBase = fechaVencimientoEfectiva(p);
  const renunciaDate = p.renuncia?.fecha ? toDate(p.renuncia.fecha) : null;
  const finTotal = renunciaDate && renunciaDate < finBase ? renunciaDate : finBase;
  const totalDays = Math.max(daysBetween(inicio, finTotal), 1);
  const segs = [];
  const p1fin = toDate(p.periodo1.fin);
  const p1end = renunciaDate && renunciaDate < p1fin ? renunciaDate : p1fin;
  segs.push({
    start: 0,
    width: (daysBetween(inicio, p1end) / totalDays) * 100,
    color: "#1B2A4A",
    title: `1er período: ${fmt(p.periodo1.inicio)} – ${fmt(p.periodo1.fin)}`,
  });
  if (p.renovacion?.inicio && !(renunciaDate && renunciaDate < toDate(p.renovacion.inicio))) {
    const rStart = daysBetween(inicio, toDate(p.renovacion.inicio));
    const rEndDate = renunciaDate && renunciaDate < toDate(p.renovacion.fin) ? renunciaDate : toDate(p.renovacion.fin);
    const rEnd = daysBetween(inicio, rEndDate);
    segs.push({
      start: (rStart / totalDays) * 100,
      width: ((rEnd - rStart) / totalDays) * 100,
      color: "#2F6E5E",
      title: `Renovación: ${fmt(p.renovacion.inicio)} – ${fmt(p.renovacion.fin)}`,
    });
  }
  const suspMarks = (p.suspensiones || [])
    .filter((s) => s.inicio && s.fin)
    .map((s) => {
      const st = daysBetween(inicio, toDate(s.inicio));
      const en = daysBetween(inicio, toDate(s.fin));
      return { left: (st / totalDays) * 100, width: Math.max(((en - st) / totalDays) * 100, 0.8), title: `Suspensión: ${fmt(s.inicio)} – ${fmt(s.fin)}${s.motivo ? " · " + s.motivo : ""}` };
    });
  return (
    <div style={{ position: "relative", height: 8, background: "#E7E9E4", borderRadius: 4, overflow: "hidden", width: "100%" }}>
      {segs.map((s, i) => (
        <div key={i} title={s.title} style={{ position: "absolute", left: `${Math.max(s.start, 0)}%`, width: `${Math.max(s.width, 0.5)}%`, top: 0, bottom: 0, background: s.color }} />
      ))}
      {suspMarks.map((s, i) => (
        <div
          key={"s" + i}
          title={s.title}
          style={{
            position: "absolute",
            left: `${s.left}%`,
            width: `${s.width}%`,
            top: 0,
            bottom: 0,
            background: "repeating-linear-gradient(45deg, #B8862F, #B8862F 3px, #F6E9C9 3px, #F6E9C9 6px)",
          }}
        />
      ))}
      {renunciaDate && (
        <div title={`Renuncia: ${fmt(p.renuncia.fecha)}`} style={{ position: "absolute", left: "100%", transform: "translateX(-2px)", top: -2, bottom: -2, width: 2, background: "#A6432D" }} />
      )}
    </div>
  );
}

// ---------- Formulario ----------
const emptyForm = () => ({
  id: null,
  dni: "",
  nombre: "",
  lugarTrabajo: "",
  habilitacionPagadora: "",
  jurisdiccion: "",
  universidad: "",
  carrera: "",
  expediente: "",
  periodo1: { inicio: "", fin: "", convenio: "" },
  renovacion: null,
  renuncia: null,
  suspensiones: [],
});

function AddToCatalog({ placeholder, onAdd }) {
  const [open, setOpen] = useState(false);
  const [val, setVal] = useState("");
  if (!open)
    return (
      <button type="button" onClick={() => setOpen(true)} style={{ background: "none", border: "none", color: "#6B5B95", cursor: "pointer", fontSize: 12, fontWeight: 600, padding: "6px 0", display: "flex", alignItems: "center", gap: 4 }}>
        <Plus size={12} /> Agregar nuevo
      </button>
    );
  return (
    <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
      <input
        autoFocus
        placeholder={placeholder}
        value={val}
        onChange={(e) => setVal(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && val.trim()) {
            onAdd(val.trim());
            setVal("");
            setOpen(false);
          }
        }}
        style={{ flex: 1, padding: "6px 9px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13 }}
      />
      <button
        type="button"
        onClick={() => {
          if (val.trim()) {
            onAdd(val.trim());
            setVal("");
          }
          setOpen(false);
        }}
        style={{ padding: "6px 10px", borderRadius: 6, border: "none", background: "#1B2A4A", color: "#fff", fontSize: 12, cursor: "pointer" }}
      >
        OK
      </button>
    </div>
  );
}

function PasantiaForm({ initial, onSave, onCancel, catalogos, onAddCatalogo, dniError }) {
  const [form, setForm] = useState(initial || emptyForm());
  const [tieneRenovacion, setTieneRenovacion] = useState(!!initial?.renovacion);
  const [tieneRenuncia, setTieneRenuncia] = useState(!!initial?.renuncia);

  const set = (path, value) => {
    setForm((f) => {
      const copy = JSON.parse(JSON.stringify(f));
      const keys = path.split(".");
      let obj = copy;
      for (let i = 0; i < keys.length - 1; i++) obj = obj[keys[i]];
      obj[keys[keys.length - 1]] = value;
      return copy;
    });
  };

  const addSuspension = () =>
    setForm((f) => ({ ...f, suspensiones: [...(f.suspensiones || []), { id: uid(), inicio: "", fin: "", motivo: "", convenio: "" }] }));
  const updateSuspension = (id, key, value) =>
    setForm((f) => ({ ...f, suspensiones: f.suspensiones.map((s) => (s.id === id ? { ...s, [key]: value } : s)) }));
  const removeSuspension = (id) => setForm((f) => ({ ...f, suspensiones: f.suspensiones.filter((s) => s.id !== id) }));

  const suggestRenovacionInicio = () => {
    if (form.periodo1.fin) {
      const d = addDays(toDate(form.periodo1.fin), 1);
      return d.toISOString().slice(0, 10);
    }
    return "";
  };

  const submit = () => {
    if (!form.dni || !form.nombre || !form.periodo1.inicio || !form.periodo1.fin) return;
    const out = { ...form, id: form.dni };
    out.renovacion = tieneRenovacion ? out.renovacion || { inicio: suggestRenovacionInicio(), fin: "", convenio: "" } : null;
    out.renuncia = tieneRenuncia ? out.renuncia || { fecha: "", nota: "", observaciones: "" } : null;
    onSave(out);
  };

  const inputStyle = {
    width: "100%",
    padding: "9px 11px",
    border: "1px solid #DADCD6",
    borderRadius: 6,
    fontSize: 14,
    fontFamily: "'IBM Plex Sans', sans-serif",
    color: "#20241F",
    background: "#fff",
  };
  const label = { fontSize: 12, fontWeight: 600, color: "#5B6158", marginBottom: 4, display: "block", textTransform: "uppercase", letterSpacing: "0.03em" };
  const section = { marginBottom: 22 };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 4 }}>
        <div>
          <label style={label}>DNI (identificador único)</label>
          <input style={{ ...inputStyle, borderColor: dniError ? "#A6432D" : "#DADCD6" }} value={form.dni} onChange={(e) => set("dni", e.target.value.replace(/[^0-9]/g, ""))} placeholder="Sin puntos" />
          {dniError && <div style={{ fontSize: 11.5, color: "#A6432D", marginTop: 3 }}>{dniError}</div>}
        </div>
        <div>
          <label style={label}>Nombre del pasante</label>
          <input style={inputStyle} value={form.nombre} onChange={(e) => set("nombre", e.target.value)} placeholder="Apellido, Nombre" />
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 22, marginTop: 18 }}>
        <div>
          <label style={label}>Lugar de trabajo</label>
          <input style={inputStyle} value={form.lugarTrabajo} onChange={(e) => set("lugarTrabajo", e.target.value)} />
        </div>
        <div>
          <label style={label}>Habilitación pagadora</label>
          <select style={inputStyle} value={form.habilitacionPagadora} onChange={(e) => set("habilitacionPagadora", e.target.value)}>
            <option value="">Seleccionar…</option>
            {catalogos.habilitaciones.map((h) => <option key={h} value={h}>{h}</option>)}
          </select>
          <AddToCatalog placeholder="Nueva habilitación pagadora" onAdd={(v) => { onAddCatalogo("habilitaciones", v); set("habilitacionPagadora", v); }} />
        </div>
        <div>
          <label style={label}>Jurisdicción</label>
          <select style={inputStyle} value={form.jurisdiccion} onChange={(e) => set("jurisdiccion", e.target.value)}>
            <option value="">Seleccionar…</option>
            {catalogos.jurisdicciones.map((j) => <option key={j} value={j}>{j}</option>)}
          </select>
          <AddToCatalog placeholder="Ej: 07 - Nueva jurisdicción" onAdd={(v) => { onAddCatalogo("jurisdicciones", v); set("jurisdiccion", v); }} />
        </div>
        <div>
          <label style={label}>Universidad / Instituto</label>
          <select style={inputStyle} value={form.universidad} onChange={(e) => set("universidad", e.target.value)}>
            <option value="">Seleccionar…</option>
            {catalogos.universidades.map((u) => (
              <option key={u.nombre} value={u.nombre}>
                {u.nombre} ({u.gastoAdmin}%{u.centralizada ? " · centralizada" : ""})
              </option>
            ))}
          </select>
          <div style={{ fontSize: 11, color: "#8A9088", marginTop: 5 }}>Las universidades se cargan desde la pestaña Catálogos (con su % de gastos administrativos).</div>
        </div>
        <div>
          <label style={label}>Expediente TIMBO</label>
          <input style={inputStyle} value={form.expediente} onChange={(e) => set("expediente", e.target.value)} />
        </div>
        <div>
          <label style={label}>Carrera</label>
          <select style={inputStyle} value={form.carrera} onChange={(e) => set("carrera", e.target.value)}>
            <option value="">Seleccionar…</option>
            {catalogos.carreras.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
          <AddToCatalog placeholder="Nueva carrera" onAdd={(v) => { onAddCatalogo("carreras", v); set("carrera", v); }} />
        </div>
      </div>

      <div style={{ ...section, padding: 14, background: "#F6F7F5", borderRadius: 8, border: "1px solid #E7E9E4" }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: "#1B2A4A", marginBottom: 10, fontFamily: "'IBM Plex Serif', serif" }}>1er período — convenio original</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 }}>
          <div>
            <label style={label}>Desde</label>
            <input type="date" style={inputStyle} value={form.periodo1.inicio} onChange={(e) => set("periodo1.inicio", e.target.value)} />
          </div>
          <div>
            <label style={label}>Hasta</label>
            <input type="date" style={inputStyle} value={form.periodo1.fin} onChange={(e) => set("periodo1.fin", e.target.value)} />
          </div>
          <div>
            <label style={label}>N° de convenio</label>
            <input style={inputStyle} value={form.periodo1.convenio || ""} onChange={(e) => set("periodo1.convenio", e.target.value)} />
          </div>
        </div>
      </div>

      <div style={{ ...section, padding: 14, background: tieneRenovacion ? "#EAF3EF" : "#F6F7F5", borderRadius: 8, border: "1px solid #E7E9E4" }}>
        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 700, color: "#1B2A4A", cursor: "pointer", fontFamily: "'IBM Plex Serif', serif" }}>
          <input
            type="checkbox"
            checked={tieneRenovacion}
            onChange={(e) => {
              setTieneRenovacion(e.target.checked);
              if (e.target.checked && !form.renovacion) set("renovacion", { inicio: suggestRenovacionInicio(), fin: "", convenio: "" });
              if (!e.target.checked) set("renovacion", null);
            }}
          />
          Tiene renovación
        </label>
        {tieneRenovacion && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14, marginTop: 10 }}>
            <div>
              <label style={label}>Desde</label>
              <input type="date" style={inputStyle} value={form.renovacion?.inicio || ""} onChange={(e) => set("renovacion.inicio", e.target.value)} />
            </div>
            <div>
              <label style={label}>Hasta</label>
              <input type="date" style={inputStyle} value={form.renovacion?.fin || ""} onChange={(e) => set("renovacion.fin", e.target.value)} />
            </div>
            <div>
              <label style={label}>N° de convenio</label>
              <input style={inputStyle} value={form.renovacion?.convenio || ""} onChange={(e) => set("renovacion.convenio", e.target.value)} />
            </div>
          </div>
        )}
      </div>

      <div style={{ ...section, padding: 14, background: "#F6F7F5", borderRadius: 8, border: "1px solid #E7E9E4" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: tieneRenuncia || form.suspensiones.length ? 10 : 0 }}>
          <span style={{ fontSize: 13, fontWeight: 700, color: "#1B2A4A", fontFamily: "'IBM Plex Serif', serif" }}>Casos especiales</span>
        </div>
        <div style={{ display: "flex", gap: 18, marginBottom: form.suspensiones.length ? 10 : 0 }}>
          <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, cursor: "pointer" }}>
            <input
              type="checkbox"
              checked={tieneRenuncia}
              onChange={(e) => {
                setTieneRenuncia(e.target.checked);
                if (e.target.checked && !form.renuncia) set("renuncia", { fecha: "", nota: "", observaciones: "" });
                if (!e.target.checked) set("renuncia", null);
              }}
            />
            Renuncia
          </label>
          <button onClick={addSuspension} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: "#6B5B95", background: "none", border: "none", cursor: "pointer", fontWeight: 600 }}>
            <Plus size={14} /> Agregar suspensión de plazo
          </button>
        </div>
        {tieneRenuncia && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14, marginBottom: form.suspensiones.length ? 10 : 0 }}>
            <div>
              <label style={label}>Fecha de renuncia</label>
              <input type="date" style={inputStyle} value={form.renuncia?.fecha || ""} onChange={(e) => set("renuncia.fecha", e.target.value)} />
            </div>
            <div>
              <label style={label}>Nota / referencia</label>
              <input style={inputStyle} value={form.renuncia?.nota || ""} onChange={(e) => set("renuncia.nota", e.target.value)} placeholder="N.º de nota fehaciente" />
            </div>
            <div>
              <label style={label}>Observaciones</label>
              <input style={inputStyle} value={form.renuncia?.observaciones || ""} onChange={(e) => set("renuncia.observaciones", e.target.value)} />
            </div>
          </div>
        )}
        {form.suspensiones.map((s) => (
          <div key={s.id} style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr auto", gap: 10, alignItems: "end", marginBottom: 8, padding: 10, background: "#fff", borderRadius: 6, border: "1px solid #E7E9E4" }}>
            <div>
              <label style={label}>Desde</label>
              <input type="date" style={inputStyle} value={s.inicio} onChange={(e) => updateSuspension(s.id, "inicio", e.target.value)} />
            </div>
            <div>
              <label style={label}>Hasta</label>
              <input type="date" style={inputStyle} value={s.fin} onChange={(e) => updateSuspension(s.id, "fin", e.target.value)} />
            </div>
            <div>
              <label style={label}>Motivo</label>
              <input style={inputStyle} value={s.motivo} onChange={(e) => updateSuspension(s.id, "motivo", e.target.value)} />
            </div>
            <div>
              <label style={label}>N° de convenio (adenda)</label>
              <input style={inputStyle} value={s.convenio || ""} onChange={(e) => updateSuspension(s.id, "convenio", e.target.value)} />
            </div>
            <button onClick={() => removeSuspension(s.id)} style={{ background: "none", border: "none", color: "#A6432D", cursor: "pointer", padding: 8 }}>
              <Trash2 size={16} />
            </button>
          </div>
        ))}
        {form.suspensiones.length > 0 && form.periodo1.inicio && form.periodo1.fin && (() => {
          const previewP = { periodo1: form.periodo1, renovacion: tieneRenovacion ? form.renovacion : null, suspensiones: form.suspensiones };
          const diasSusp = diasSuspendidos(previewP);
          if (!diasSusp) return null;
          const nuevaFecha = fechaVencimientoEfectiva(previewP);
          return (
            <div style={{ marginTop: 6, padding: "9px 12px", background: "#EAF3EF", borderRadius: 6, fontSize: 12.5, color: "#2F6E5E" }}>
              Con {diasSusp} día{diasSusp !== 1 ? "s" : ""} de suspensión, la fecha de finalización se corre al <strong>{fmt(nuevaFecha.toISOString().slice(0, 10))}</strong>.
            </div>
          );
        })()}
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 6 }}>
        <button onClick={onCancel} style={{ padding: "10px 18px", borderRadius: 6, border: "1px solid #DADCD6", background: "#fff", color: "#5B6158", fontSize: 14, cursor: "pointer", fontWeight: 600 }}>
          Cancelar
        </button>
        <button onClick={submit} style={{ padding: "10px 18px", borderRadius: 6, border: "none", background: "#1B2A4A", color: "#fff", fontSize: 14, cursor: "pointer", fontWeight: 600 }}>
          Guardar pasantía
        </button>
      </div>
    </div>
  );
}

// ---------- Aviso de vencimiento imprimible (→ PDF vía impresión del navegador) ----------
function textoAviso(p, estado) {
  const efectiva = fmt(fechaVencimientoEfectiva(p).toISOString().slice(0, 10));
  if (estado === "por_vencer_renovacion") {
    return `Se informa que la pasantía de ${p.nombre} (DNI ${p.dni}), correspondiente al 1er período de contratación, vence el ${efectiva}. De corresponder su continuidad, deberá iniciarse la gestión de renovación de convenio dentro del expediente${p.expediente ? ` N.º ${p.expediente}` : ""}, con la debida antelación.`;
  }
  return `Se informa que la pasantía de ${p.nombre} (DNI ${p.dni}) finaliza de manera definitiva el ${efectiva}, no correspondiendo nuevas renovaciones. Se solicita tomar las previsiones administrativas correspondientes.`;
}

function DatosAviso({ p }) {
  const row = { display: "flex", justifyContent: "space-between", padding: "7px 0", borderBottom: "1px solid #E7E9E4", fontSize: 13 };
  const k = { color: "#5B6158" };
  const v = { fontWeight: 600 };
  return (
    <div style={{ marginBottom: 20 }}>
      {[
        ["Nombre", p.nombre],
        ["DNI", p.dni],
        ["Lugar de trabajo", p.lugarTrabajo || "—"],
        ["Jurisdicción", p.jurisdiccion || "—"],
        ["Habilitación pagadora", p.habilitacionPagadora || "—"],
        ["Universidad / Instituto", p.universidad || "—"],
        ["Expediente TIMBO", p.expediente || "—"],
      ].map(([kk, vv]) => (
        <div key={kk} style={row}>
          <span style={k}>{kk}</span>
          <span style={v}>{vv}</span>
        </div>
      ))}
    </div>
  );
}

function AvisoPDF({ items, onClose }) {
  useEffect(() => {
    const t = setTimeout(() => window.print(), 150);
    return () => clearTimeout(t);
  }, []);
  const fechaHoy = new Date().toLocaleDateString("es-AR", { day: "2-digit", month: "long", year: "numeric" });
  const single = items.length === 1;

  return (
    <div
      className="print-area"
      style={{ position: "fixed", inset: 0, background: "#fff", zIndex: 100, overflowY: "auto", padding: "44px 56px", fontFamily: "'IBM Plex Sans', sans-serif", color: "#20241F" }}
    >
      <button
        className="no-print"
        onClick={onClose}
        style={{ position: "fixed", top: 20, right: 26, padding: "9px 16px", borderRadius: 7, border: "1px solid #DADCD6", background: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 600 }}
      >
        Cerrar
      </button>

      <div style={{ borderBottom: "3px solid #1B2A4A", paddingBottom: 16, marginBottom: 26, display: "flex", justifyContent: "space-between", alignItems: "flex-end" }}>
        <div>
          <div style={{ fontFamily: "'IBM Plex Serif', serif", fontSize: 21, fontWeight: 700, color: "#1B2A4A" }}>
            {single ? "Aviso de vencimiento de pasantía" : "Aviso de vencimientos de pasantías"}
          </div>
          <div style={{ fontSize: 12.5, color: "#5B6158", marginTop: 4 }}>Generado el {fechaHoy}</div>
        </div>
      </div>

      {single ? (
        <div>
          <DatosAviso p={items[0].p} />
          <div style={{ fontSize: 13.5, lineHeight: 1.7, textAlign: "justify" }}>{textoAviso(items[0].p, items[0].estado)}</div>
          <div style={{ marginTop: 60, display: "flex", justifyContent: "space-between" }}>
            <div style={{ textAlign: "center", width: 220 }}>
              <div style={{ borderTop: "1px solid #20241F", paddingTop: 6, fontSize: 12, color: "#5B6158" }}>Firma responsable</div>
            </div>
            <div style={{ textAlign: "center", width: 220 }}>
              <div style={{ borderTop: "1px solid #20241F", paddingTop: 6, fontSize: 12, color: "#5B6158" }}>Fecha de notificación</div>
            </div>
          </div>
        </div>
      ) : (
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
          <thead>
            <tr style={{ background: "#F6F7F5" }}>
              {["Nombre", "DNI", "Universidad/Instituto", "Jurisdicción", "Tipo de aviso", "Vence", "Días restantes"].map((h) => (
                <th key={h} style={{ textAlign: "left", padding: "9px 8px", borderBottom: "2px solid #1B2A4A", fontWeight: 700, color: "#1B2A4A" }}>
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {items.map(({ p, estado, restantes }) => (
              <tr key={p.id} style={{ borderBottom: "1px solid #E7E9E4" }}>
                <td style={{ padding: "8px" }}>{p.nombre}</td>
                <td style={{ padding: "8px" }}>{p.dni}</td>
                <td style={{ padding: "8px" }}>{p.universidad}</td>
                <td style={{ padding: "8px" }}>{p.jurisdiccion}</td>
                <td style={{ padding: "8px" }}>{estado === "por_vencer_renovacion" ? "Iniciar renovación" : "Finalización definitiva"}</td>
                <td style={{ padding: "8px" }}>{fmt(fechaVencimientoEfectiva(p).toISOString().slice(0, 10))}</td>
                <td style={{ padding: "8px", fontWeight: 700 }}>{restantes}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

function normalizeCatalogos(raw) {
  const merged = { ...DEFAULT_CATALOGOS, ...(raw || {}) };
  merged.universidades = (merged.universidades || []).map((u) =>
    typeof u === "string" ? { nombre: u, gastoAdmin: 0, centralizada: false } : u
  );
  merged.jurisdicciones = merged.jurisdicciones || [];
  merged.habilitaciones = merged.habilitaciones || [];
  merged.carreras = merged.carreras || [];
  merged.estadosTramite = merged.estadosTramite || [];
  // Compatibilidad con backups viejos donde asignaciones no tenían id/creadoEl/norma.
  merged.asignaciones = (merged.asignaciones || []).map((a) => ({
    id: a.id || uid(),
    year: a.year,
    month: a.month,
    monto: a.monto,
    norma: a.norma || "",
    observacion: a.observacion || "",
    creadoEl: a.creadoEl || "2000-01-01T00:00:00.000Z",
  }));
  merged.centralizadaNombre = merged.centralizadaNombre || "";
  merged.seguridad = { pwdEdicion: "", pwdLiquidacion: "", pwdCatalogo: "", ...(merged.seguridad || {}) };
  return merged;
}

const emptyTramite = () => ({
  id: null,
  expediente: "",
  lugarTrabajo: "",
  habilitacionPagadora: "",
  jurisdiccion: "",
  estado: "",
  observaciones: "",
  cantidadMeses: "",
  fechaProbableInicio: "",
  detalles: [{ id: uid(), carrera: "", universidad: "", cantidadPasantes: "" }],
});

function TramiteForm({ initial, onSave, onCancel, catalogos, onAddCatalogo }) {
  const [form, setForm] = useState(initial || emptyTramite());
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));

  const inputStyle = {
    width: "100%",
    padding: "9px 11px",
    border: "1px solid #DADCD6",
    borderRadius: 6,
    fontSize: 14,
    fontFamily: "'IBM Plex Sans', sans-serif",
    color: "#20241F",
    background: "#fff",
  };
  const label = { fontSize: 12, fontWeight: 600, color: "#5B6158", marginBottom: 4, display: "block", textTransform: "uppercase", letterSpacing: "0.03em" };

  const totalPasantes = (form.detalles || []).reduce((acc, d) => acc + (Number(d.cantidadPasantes) || 0), 0);

  const updateDetalle = (id, key, value) =>
    setForm((f) => ({ ...f, detalles: f.detalles.map((d) => (d.id === id ? { ...d, [key]: value } : d)) }));
  const addDetalle = () =>
    setForm((f) => ({ ...f, detalles: [...f.detalles, { id: uid(), carrera: "", universidad: "", cantidadPasantes: "" }] }));
  const removeDetalle = (id) =>
    setForm((f) => ({ ...f, detalles: f.detalles.length > 1 ? f.detalles.filter((d) => d.id !== id) : f.detalles }));

  const submit = () => {
    if (!form.expediente) return;
    onSave({ ...form, id: form.id || uid() });
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
        <div>
          <label style={label}>Expediente</label>
          <input style={inputStyle} value={form.expediente} onChange={(e) => set("expediente", e.target.value)} />
        </div>
        <div>
          <label style={label}>Cantidad de meses a contratar</label>
          <input type="number" min="1" style={inputStyle} value={form.cantidadMeses} onChange={(e) => set("cantidadMeses", e.target.value)} />
        </div>
        <div>
          <label style={label}>Fecha probable de inicio</label>
          <input type="date" style={inputStyle} value={form.fechaProbableInicio} onChange={(e) => set("fechaProbableInicio", e.target.value)} />
        </div>
        <div>
          <label style={label}>Lugar de trabajo</label>
          <input style={inputStyle} value={form.lugarTrabajo} onChange={(e) => set("lugarTrabajo", e.target.value)} />
        </div>
        <div>
          <label style={label}>Habilitación pagadora</label>
          <select style={inputStyle} value={form.habilitacionPagadora} onChange={(e) => set("habilitacionPagadora", e.target.value)}>
            <option value="">Seleccionar…</option>
            {catalogos.habilitaciones.map((h) => <option key={h} value={h}>{h}</option>)}
          </select>
        </div>
        <div>
          <label style={label}>Jurisdicción</label>
          <select style={inputStyle} value={form.jurisdiccion} onChange={(e) => set("jurisdiccion", e.target.value)}>
            <option value="">Seleccionar…</option>
            {catalogos.jurisdicciones.map((j) => <option key={j} value={j}>{j}</option>)}
          </select>
        </div>
        <div>
          <label style={label}>Estado del trámite</label>
          <select style={inputStyle} value={form.estado} onChange={(e) => set("estado", e.target.value)}>
            <option value="">Seleccionar…</option>
            {catalogos.estadosTramite.map((e) => <option key={e} value={e}>{e}</option>)}
          </select>
          <AddToCatalog placeholder="Nuevo estado (ej: En revisión legal)" onAdd={(v) => { onAddCatalogo("estadosTramite", v); set("estado", v); }} />
        </div>
      </div>

      <div style={{ padding: 14, background: "#F6F7F5", borderRadius: 8, border: "1px solid #E7E9E4" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: "#1B2A4A", fontFamily: "'IBM Plex Serif', serif" }}>Carreras / universidades solicitadas</div>
          <div style={{ fontSize: 12.5, color: "#5B6158" }}>Total: <strong>{totalPasantes}</strong> pasante{totalPasantes !== 1 ? "s" : ""}</div>
        </div>
        {(form.detalles || []).map((d) => (
          <div key={d.id} style={{ display: "grid", gridTemplateColumns: "1.4fr 1.4fr 100px auto", gap: 10, alignItems: "end", marginBottom: 8, padding: 10, background: "#fff", borderRadius: 6, border: "1px solid #E7E9E4" }}>
            <div>
              <label style={label}>Carrera</label>
              <select style={inputStyle} value={d.carrera} onChange={(e) => updateDetalle(d.id, "carrera", e.target.value)}>
                <option value="">Seleccionar…</option>
                {catalogos.carreras.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
              <AddToCatalog placeholder="Nueva carrera" onAdd={(v) => { onAddCatalogo("carreras", v); updateDetalle(d.id, "carrera", v); }} />
            </div>
            <div>
              <label style={label}>Universidad / Instituto</label>
              <select style={inputStyle} value={d.universidad} onChange={(e) => updateDetalle(d.id, "universidad", e.target.value)}>
                <option value="">Seleccionar…</option>
                {catalogos.universidades.map((u) => <option key={u.nombre} value={u.nombre}>{u.nombre}</option>)}
              </select>
            </div>
            <div>
              <label style={label}>Cantidad</label>
              <input type="number" min="1" style={inputStyle} value={d.cantidadPasantes} onChange={(e) => updateDetalle(d.id, "cantidadPasantes", e.target.value)} />
            </div>
            <button onClick={() => removeDetalle(d.id)} disabled={form.detalles.length === 1} style={{ background: "none", border: "none", color: form.detalles.length === 1 ? "#C7CDDB" : "#A6432D", cursor: form.detalles.length === 1 ? "not-allowed" : "pointer", padding: 8 }}>
              <Trash2 size={16} />
            </button>
          </div>
        ))}
        <button type="button" onClick={addDetalle} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", color: "#6B5B95", cursor: "pointer", fontSize: 13, fontWeight: 600, padding: "4px 0" }}>
          <Plus size={14} /> Agregar otra carrera / universidad
        </button>
      </div>

      <div>
        <label style={label}>Observaciones</label>
        <textarea style={{ ...inputStyle, minHeight: 70, resize: "vertical" }} value={form.observaciones} onChange={(e) => set("observaciones", e.target.value)} />
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 6 }}>
        <button onClick={onCancel} style={{ padding: "10px 18px", borderRadius: 6, border: "1px solid #DADCD6", background: "#fff", color: "#5B6158", fontSize: 14, cursor: "pointer", fontWeight: 600 }}>
          Cancelar
        </button>
        <button onClick={submit} style={{ padding: "10px 18px", borderRadius: 6, border: "none", background: "#1B2A4A", color: "#fff", fontSize: 14, cursor: "pointer", fontWeight: 600 }}>
          Guardar trámite
        </button>
      </div>
    </div>
  );
}

function LiquidacionView({ liq, etiqueta, nombreCentralizada, onExport, onPrint, eeMap, onAsignarEE, readOnly, onConfirmar, confirmado, fechaConfirmacion }) {
  const th = { textAlign: "left", padding: "7px 8px", fontSize: 11.5, fontWeight: 700, color: "#5B6158", textTransform: "uppercase", letterSpacing: "0.02em", borderBottom: "2px solid #E7E9E4" };
  const td = { padding: "7px 8px", fontSize: 12.5, borderBottom: "1px solid #F0F1EE" };
  const [seleccionados, setSeleccionados] = useState([]);
  const [eeInput, setEeInput] = useState("");
  const [printScope, setPrintScope] = useState("todo");
  const [printValue, setPrintValue] = useState("");
  const [filtroJurisdiccion, setFiltroJurisdiccion] = useState("");
  const [soloCentralizada, setSoloCentralizada] = useState(false);

  const totalGeneral = liq.porHabilitacion.reduce((acc, g) => acc + g.total, 0);
  const jurisdiccionesDisponibles = Array.from(new Set(liq.filas.map((f) => f.p.jurisdiccion).filter(Boolean))).sort();

  const toggleSel = (dni) => setSeleccionados((s) => (s.includes(dni) ? s.filter((x) => x !== dni) : [...s, dni]));
  const asignar = () => {
    if (!eeInput.trim() || seleccionados.length === 0) return;
    onAsignarEE(seleccionados, eeInput.trim());
    setSeleccionados([]);
    setEeInput("");
  };

  const ejecutarImprimir = () => {
    if (printScope === "todo") onPrint(null);
    else if (printScope === "habilitacion") onPrint({ tipo: "habilitacion", valor: printValue || liq.porHabilitacion[0]?.habilitacion });
    else if (printScope === "jurisdiccion") onPrint({ tipo: "jurisdiccion", valor: printValue || liq.centralizada[0]?.jurisdiccion });
    else if (printScope === "centralizada") onPrint({ tipo: "centralizada" });
  };

  // Vista filtrada en pantalla (no afecta Excel/impresión, que tienen su propio alcance)
  const porHabilitacionVista = soloCentralizada
    ? []
    : liq.porHabilitacion
        .map((g) => {
          const filas = filtroJurisdiccion ? g.filas.filter((f) => f.p.jurisdiccion === filtroJurisdiccion) : g.filas;
          if (!filtroJurisdiccion) return g;
          // Recalcular subtotales solo con las filas visibles, para que no mezclen otras jurisdicciones.
          const subtotalAsigIapos = filas.reduce((acc, f) => acc + f.monto + f.iapos, 0);
          const gastoAdminTotal = filas.reduce((acc, f) => acc + (f.centralizada ? 0 : f.gastoAdmin), 0);
          return { ...g, filas, subtotalAsigIapos, gastoAdminTotal, total: subtotalAsigIapos + gastoAdminTotal };
        })
        .filter((g) => g.filas.length > 0);
  const centralizadaVista = filtroJurisdiccion ? liq.centralizada.filter((g) => g.jurisdiccion === filtroJurisdiccion) : liq.centralizada;

  // Total real de todo lo que cuesta la jurisdicción filtrada (incluye tanto lo que paga
  // la habilitación pagadora como el gasto administrativo que paga la centralizada, ya
  // que ambos son costo de esa jurisdicción, sin importar quién efectivamente lo abona).
  const totalJurisdiccionFiltrada = filtroJurisdiccion
    ? liq.filas.filter((f) => f.p.jurisdiccion === filtroJurisdiccion).reduce((acc, f) => acc + f.total, 0)
    : null;

  const periodoLabel = (f) => (f.periodoTipo === "renovacion" ? "Renovación" : "Período inicial");

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 10 }}>
        <div style={{ fontSize: 13, color: "#5B6158" }}>
          {liq.filas.length} pasantía{liq.filas.length !== 1 ? "s" : ""} liquidada{liq.filas.length !== 1 ? "s" : ""} · Total general{" "}
          <strong style={{ color: "#1B2A4A" }}>{moneyFmt(totalGeneral)}</strong>
          {confirmado && <span style={{ marginLeft: 10, padding: "2px 9px", borderRadius: 20, background: "#EAF3EF", color: "#2F6E5E", fontSize: 11.5, fontWeight: 600 }}>Confirmada el {fmt(fechaConfirmacion?.slice(0, 10))}</span>}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={onExport} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: "#fff", color: "#1B2A4A", border: "1px solid #DADCD6", borderRadius: 7, fontSize: 13.5, fontWeight: 600, cursor: "pointer" }}>
            <Download size={14} /> Excel
          </button>
          {onConfirmar && !confirmado && (
            <button onClick={onConfirmar} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: "#2F6E5E", color: "#fff", border: "none", borderRadius: 7, fontSize: 13.5, fontWeight: 600, cursor: "pointer" }}>
              <CheckCircle2 size={14} /> Confirmar liquidación
            </button>
          )}
        </div>
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12, padding: "8px 12px", background: "#F6F7F5", borderRadius: 8, fontSize: 12.5, flexWrap: "wrap" }}>
        <span style={{ color: "#5B6158" }}>Ver:</span>
        <select value={filtroJurisdiccion} onChange={(e) => setFiltroJurisdiccion(e.target.value)} style={{ padding: "6px 8px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 12.5 }}>
          <option value="">Todas las jurisdicciones</option>
          {jurisdiccionesDisponibles.map((j) => <option key={j} value={j}>{j}</option>)}
        </select>
        {liq.centralizada.length > 0 && (
          <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
            <input type="checkbox" checked={soloCentralizada} onChange={(e) => setSoloCentralizada(e.target.checked)} />
            Solo habilitación centralizada
          </label>
        )}
        {filtroJurisdiccion && (
          <span style={{ marginLeft: "auto", fontWeight: 700, color: "#1B2A4A" }}>
            Total a liquidar en {filtroJurisdiccion}: {moneyFmt(totalJurisdiccionFiltrada)}
          </span>
        )}
      </div>

      {onPrint && (
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12, padding: "8px 12px", background: "#F6F7F5", borderRadius: 8, fontSize: 12.5, flexWrap: "wrap" }}>
          <Printer size={14} color="#5B6158" />
          <span style={{ color: "#5B6158" }}>Imprimir / PDF:</span>
          <select value={printScope} onChange={(e) => { setPrintScope(e.target.value); setPrintValue(""); }} style={{ padding: "6px 8px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 12.5 }}>
            <option value="todo">Todo</option>
            <option value="habilitacion">Por habilitación pagadora</option>
            {liq.centralizada.length > 0 && <option value="jurisdiccion">Centralizada — por jurisdicción</option>}
            {liq.centralizada.length > 0 && <option value="centralizada">Centralizada — todas las jurisdicciones</option>}
          </select>
          {printScope === "habilitacion" && (
            <select value={printValue} onChange={(e) => setPrintValue(e.target.value)} style={{ padding: "6px 8px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 12.5 }}>
              {liq.porHabilitacion.map((g) => <option key={g.habilitacion} value={g.habilitacion}>{g.habilitacion}</option>)}
            </select>
          )}
          {printScope === "jurisdiccion" && (
            <select value={printValue} onChange={(e) => setPrintValue(e.target.value)} style={{ padding: "6px 8px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 12.5 }}>
              {liq.centralizada.map((g) => <option key={g.jurisdiccion} value={g.jurisdiccion}>{g.jurisdiccion}</option>)}
            </select>
          )}
          <button onClick={ejecutarImprimir} style={{ padding: "6px 12px", borderRadius: 6, border: "none", background: "#1B2A4A", color: "#fff", fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}>
            Generar
          </button>
        </div>
      )}

      {!readOnly && (
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16, padding: "8px 12px", background: seleccionados.length ? "#EAF3EF" : "#F6F7F5", borderRadius: 8, fontSize: 12.5 }}>
          <span style={{ color: "#5B6158" }}>{seleccionados.length} seleccionado{seleccionados.length !== 1 ? "s" : ""}</span>
          <input
            placeholder="N° de EE"
            value={eeInput}
            onChange={(e) => setEeInput(e.target.value)}
            disabled={seleccionados.length === 0}
            style={{ padding: "6px 9px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 12.5, width: 160 }}
          />
          <button onClick={asignar} disabled={seleccionados.length === 0 || !eeInput.trim()} style={{ padding: "6px 12px", borderRadius: 6, border: "none", background: seleccionados.length && eeInput.trim() ? "#1B2A4A" : "#C7CDDB", color: "#fff", fontSize: 12.5, fontWeight: 600, cursor: seleccionados.length ? "pointer" : "not-allowed" }}>
            Asignar EE a seleccionados
          </button>
          {seleccionados.length > 0 && (
            <button onClick={() => setSeleccionados([])} style={{ background: "none", border: "none", color: "#8A9088", fontSize: 12, cursor: "pointer" }}>
              Limpiar selección
            </button>
          )}
        </div>
      )}

      {porHabilitacionVista.length === 0 && (
        <div style={{ background: "#fff", border: "1px dashed #DADCD6", borderRadius: 10, padding: 30, textAlign: "center", color: "#8A9088" }}>
          No hay pasantías para mostrar con este filtro.
        </div>
      )}

      {porHabilitacionVista.map((g) => (
        <div key={g.habilitacion} style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, marginBottom: 14, overflow: "hidden" }}>
          <div style={{ padding: "11px 16px", background: "#F6F7F5", borderBottom: "1px solid #E7E9E4", fontWeight: 700, fontSize: 14, color: "#1B2A4A", fontFamily: "'IBM Plex Serif', serif" }}>
            {g.habilitacion}
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  {!readOnly && <th style={th}></th>}
                  <th style={th}>Pasante</th>
                  <th style={th}>Universidad</th>
                  <th style={th}>Período</th>
                  <th style={th}>Vigencia</th>
                  <th style={th}>Días</th>
                  <th style={th}>Asignación</th>
                  <th style={th}>IAPOS</th>
                  <th style={th}>% Adm.</th>
                  <th style={th}>Gasto adm.</th>
                  <th style={th}>Total</th>
                  <th style={th}>N° EE</th>
                </tr>
              </thead>
              <tbody>
                {g.filas.map((f) => (
                  <tr key={f.p.id} style={f.centralizada ? { background: "#FBF3E4" } : undefined}>
                    {!readOnly && (
                      <td style={td}>
                        <input type="checkbox" checked={seleccionados.includes(f.p.dni)} onChange={() => toggleSel(f.p.dni)} />
                      </td>
                    )}
                    <td style={td}>
                      {f.p.nombre}
                      {f.motivoDiff && (
                        <div style={{ fontSize: 10.5, color: f.motivoDiff === "ya no vigente" ? "#A6432D" : "#2F6E5E", fontWeight: 600, marginTop: 2 }}>
                          {f.motivoDiff === "alta nueva" ? "Alta no incluida antes" : f.motivoDiff === "ya no vigente" ? "Ya no vigente / baja" : "Ajuste"}
                        </div>
                      )}
                      {f.centralizada && (
                        <div style={{ fontSize: 10.5, color: "#B8862F", fontWeight: 600, marginTop: 2 }}>
                          Gasto administrativo se abona por {nombreCentralizada || "la habilitación centralizada"}
                        </div>
                      )}
                    </td>
                    <td style={td}>{f.p.universidad}</td>
                    <td style={td}>{periodoLabel(f)}</td>
                    <td style={td}>{fmt(f.vigenciaDesde)} – {fmt(f.vigenciaHasta)}</td>
                    <td style={td}>{f.dias}{f.esMesCompleto ? " (completo)" : ""}</td>
                    <td style={td}>{moneyFmt(f.monto)}</td>
                    <td style={td}>{moneyFmt(f.iapos)}</td>
                    <td style={td}>{f.gastoPct}%</td>
                    <td style={td}>{f.centralizada ? "—" : moneyFmt(f.gastoAdmin)}</td>
                    <td style={{ ...td, fontWeight: 700 }}>{moneyFmt(totalACargoHabilitacion(f))}</td>
                    <td style={td}>{(eeMap && eeMap[f.p.dni]) || "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ padding: "10px 16px", background: "#F6F7F5", borderTop: "1px solid #E7E9E4", display: "flex", justifyContent: "flex-end", gap: 24, fontSize: 12.5 }}>
            <span>Subtotal (asig.+IAPOS): <strong>{moneyFmt(g.subtotalAsigIapos)}</strong></span>
            <span>Gasto admin.: <strong>{moneyFmt(g.gastoAdminTotal)}</strong></span>
            <span style={{ color: "#1B2A4A" }}>Total a pagar: <strong>{moneyFmt(g.total)}</strong></span>
          </div>
        </div>
      ))}

      {centralizadaVista.length > 0 && (
        <div style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, marginTop: 22, overflow: "hidden" }}>
          <div style={{ padding: "12px 16px", background: "#1B2A4A", color: "#fff", fontWeight: 700, fontSize: 14.5, fontFamily: "'IBM Plex Serif', serif" }}>
            Liquidación centralizada{nombreCentralizada ? ` — ${nombreCentralizada}` : ""}
          </div>
          <div style={{ padding: "10px 16px", fontSize: 12, color: "#8A9088", borderBottom: "1px solid #E7E9E4" }}>
            Detalle del gasto administrativo de estas pasantías (universidades centralizadas), agrupado por jurisdicción, para que la habilitación centralizada realice ese pago. La asignación estímulo y el IAPOS de estas pasantías ya están incluidos en la liquidación de su habilitación pagadora, arriba.
          </div>
          {centralizadaVista.map((g) => (
            <div key={g.jurisdiccion}>
              <div style={{ padding: "9px 16px", background: "#F6F7F5", fontWeight: 600, fontSize: 13, color: "#1B2A4A" }}>{g.jurisdiccion}</div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr>
                      {!readOnly && <th style={th}></th>}
                      <th style={th}>Pasante</th>
                      <th style={th}>Habilitación pagadora (original)</th>
                      <th style={th}>Lugar de trabajo</th>
                      <th style={th}>Período</th>
                      <th style={th}>Vigencia</th>
                      <th style={th}>% Adm.</th>
                      <th style={th}>Gasto administrativo</th>
                      <th style={th}>N° EE</th>
                    </tr>
                  </thead>
                  <tbody>
                    {g.filas.map((f) => (
                      <tr key={f.p.id}>
                        {!readOnly && (
                          <td style={td}>
                            <input type="checkbox" checked={seleccionados.includes(f.p.dni)} onChange={() => toggleSel(f.p.dni)} />
                          </td>
                        )}
                        <td style={td}>{f.p.nombre}</td>
                        <td style={td}>{f.p.habilitacionPagadora}</td>
                        <td style={td}>{f.p.lugarTrabajo}</td>
                        <td style={td}>{periodoLabel(f)}</td>
                        <td style={td}>{fmt(f.vigenciaDesde)} – {fmt(f.vigenciaHasta)}</td>
                        <td style={td}>{f.gastoPct}%</td>
                        <td style={{ ...td, fontWeight: 700 }}>{moneyFmt(f.gastoAdmin)}</td>
                        <td style={td}>{(eeMap && eeMap[f.p.dni]) || "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{ padding: "8px 16px", textAlign: "right", fontSize: 12.5, borderBottom: "1px solid #F0F1EE" }}>
                Subtotal {g.jurisdiccion}: <strong>{moneyFmt(g.subtotal)}</strong>
              </div>
            </div>
          ))}
          <div style={{ padding: "12px 16px", textAlign: "right", fontSize: 13.5, background: "#F6F7F5" }}>
            Total centralizada: <strong style={{ color: "#1B2A4A" }}>{moneyFmt(liq.centralizadaTotal)}</strong>
          </div>
        </div>
      )}
    </div>
  );
}

function LiquidacionPrint({ liq, etiqueta, nombreCentralizada, eeMap, filtro, onClose }) {
  useEffect(() => {
    const t = setTimeout(() => window.print(), 150);
    return () => clearTimeout(t);
  }, []);

  const th = { textAlign: "left", padding: "6px 7px", fontSize: 10.5, fontWeight: 700, color: "#5B6158", textTransform: "uppercase", borderBottom: "2px solid #1B2A4A" };
  const td = { padding: "6px 7px", fontSize: 11.5, borderBottom: "1px solid #E7E9E4" };

  let porHabilitacion = liq.porHabilitacion;
  let centralizada = liq.centralizada;
  let titulo = `Liquidación — ${etiqueta}`;

  if (filtro?.tipo === "habilitacion") {
    porHabilitacion = liq.porHabilitacion.filter((g) => g.habilitacion === filtro.valor);
    centralizada = [];
    titulo = `Liquidación — ${etiqueta} — ${filtro.valor}`;
  } else if (filtro?.tipo === "jurisdiccion") {
    porHabilitacion = [];
    centralizada = liq.centralizada.filter((g) => g.jurisdiccion === filtro.valor);
    titulo = `Liquidación centralizada — ${etiqueta} — ${filtro.valor}`;
  } else if (filtro?.tipo === "centralizada") {
    porHabilitacion = [];
    titulo = `Liquidación centralizada — ${etiqueta}`;
  }

  const fechaHoy = new Date().toLocaleDateString("es-AR", { day: "2-digit", month: "long", year: "numeric" });

  return (
    <div className="print-area" style={{ position: "fixed", inset: 0, background: "#fff", zIndex: 100, overflowY: "auto", padding: "36px 44px", fontFamily: "'IBM Plex Sans', sans-serif", color: "#20241F" }}>
      <button className="no-print" onClick={onClose} style={{ position: "fixed", top: 20, right: 26, padding: "9px 16px", borderRadius: 7, border: "1px solid #DADCD6", background: "#fff", cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
        Cerrar
      </button>
      <div style={{ borderBottom: "3px solid #1B2A4A", paddingBottom: 14, marginBottom: 22 }}>
        <div style={{ fontFamily: "'IBM Plex Serif', serif", fontSize: 19, fontWeight: 700, color: "#1B2A4A" }}>{titulo}</div>
        <div style={{ fontSize: 12, color: "#5B6158", marginTop: 3 }}>Generado el {fechaHoy}</div>
      </div>

      {porHabilitacion.map((g) => (
        <div key={g.habilitacion} style={{ marginBottom: 20, breakInside: "avoid" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: "#1B2A4A", marginBottom: 6 }}>{g.habilitacion}</div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={th}>Pasante</th>
                <th style={th}>Universidad</th>
                <th style={th}>Días</th>
                <th style={th}>Asignación</th>
                <th style={th}>IAPOS</th>
                <th style={th}>Gasto adm.</th>
                <th style={th}>Total</th>
                <th style={th}>N° EE</th>
              </tr>
            </thead>
            <tbody>
              {g.filas.map((f) => (
                <tr key={f.p.id}>
                  <td style={td}>{f.p.nombre}{f.centralizada ? " *" : ""}</td>
                  <td style={td}>{f.p.universidad}</td>
                  <td style={td}>{f.dias}</td>
                  <td style={td}>{moneyFmt(f.monto)}</td>
                  <td style={td}>{moneyFmt(f.iapos)}</td>
                  <td style={td}>{f.centralizada ? "—" : moneyFmt(f.gastoAdmin)}</td>
                  <td style={{ ...td, fontWeight: 700 }}>{moneyFmt(totalACargoHabilitacion(f))}</td>
                  <td style={td}>{(eeMap && eeMap[f.p.dni]) || ""}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{ textAlign: "right", fontSize: 12, marginTop: 4 }}>
            Subtotal (asig.+IAPOS): <strong>{moneyFmt(g.subtotalAsigIapos)}</strong> · Gasto adm.: <strong>{moneyFmt(g.gastoAdminTotal)}</strong> · Total: <strong>{moneyFmt(g.total)}</strong>
          </div>
          {g.filas.some((f) => f.centralizada) && <div style={{ fontSize: 10.5, color: "#8A9088", marginTop: 4 }}>* Gasto administrativo abonado por {nombreCentralizada || "la habilitación centralizada"}.</div>}
        </div>
      ))}

      {centralizada.map((g) => (
        <div key={g.jurisdiccion} style={{ marginBottom: 20, breakInside: "avoid" }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, color: "#1B2A4A", marginBottom: 6 }}>{g.jurisdiccion}{nombreCentralizada ? ` — ${nombreCentralizada}` : ""}</div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={th}>Pasante</th>
                <th style={th}>Habilitación pagadora</th>
                <th style={th}>Lugar de trabajo</th>
                <th style={th}>% Adm.</th>
                <th style={th}>Gasto administrativo</th>
              </tr>
            </thead>
            <tbody>
              {g.filas.map((f) => (
                <tr key={f.p.id}>
                  <td style={td}>{f.p.nombre}</td>
                  <td style={td}>{f.p.habilitacionPagadora}</td>
                  <td style={td}>{f.p.lugarTrabajo}</td>
                  <td style={td}>{f.gastoPct}%</td>
                  <td style={{ ...td, fontWeight: 700 }}>{moneyFmt(f.gastoAdmin)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{ textAlign: "right", fontSize: 12, marginTop: 4 }}>Subtotal: <strong>{moneyFmt(g.subtotal)}</strong></div>
        </div>
      ))}
    </div>
  );
}

function liqPeriodKey(tipo, year, month) {
  return `${tipo}-${year}-${month}`;
}

// ---------- Presupuesto estimado de trámites (contrataciones futuras) ----------
function getAsignacionParaPresupuesto(catalogos, year, month) {
  const exactoMonto = getAsignacion(catalogos, year, month);
  if (exactoMonto !== null) return { monto: exactoMonto, estimado: false };
  if (!catalogos.asignaciones || catalogos.asignaciones.length === 0) return null;
  const latest = catalogos.asignaciones.slice().sort((a, b) => b.year * 12 + b.month - (a.year * 12 + a.month))[0];
  return { monto: latest.monto, estimado: true };
}

function calcularPresupuestoTramite(t, catalogos) {
  const renglones = (t.detalles || []).filter((d) => d.universidad && Number(d.cantidadPasantes) > 0);
  if (!t.fechaProbableInicio || !t.cantidadMeses || renglones.length === 0) return null;
  const inicio = toDate(t.fechaProbableInicio);
  if (!inicio) return null;
  const year = inicio.getFullYear();
  const startMonth = inicio.getMonth();
  const startDay = inicio.getDate();
  const mesesHastaFinAno = 12 - startMonth;
  const mesesACalcular = Math.min(Number(t.cantidadMeses), mesesHastaFinAno);
  if (mesesACalcular <= 0) return null;

  const cantidadPasantesTotal = renglones.reduce((acc, r) => acc + (Number(r.cantidadPasantes) || 0), 0);
  const detalleMeses = [];
  let totalGeneral = 0;
  let faltanDatos = false;

  for (let i = 0; i < mesesACalcular; i++) {
    const month = startMonth + i;
    const asig = getAsignacionParaPresupuesto(catalogos, year, month);
    if (!asig) {
      faltanDatos = true;
      detalleMeses.push({ year, month, sinDato: true, porRenglon: [], totalMes: 0 });
      continue;
    }
    let montoBase = asig.monto;
    let esProrrateado = false;
    let diasTrabajados = null;
    if (i === 0 && startDay > 1) {
      diasTrabajados = diasEnMes(year, month) - startDay + 1;
      montoBase = (asig.monto / 30) * diasTrabajados;
      esProrrateado = true;
    }
    const iaposBase = montoBase * (IAPOS_PCT / 100);

    const porRenglon = renglones.map((r) => {
      const uni = (catalogos.universidades || []).find((u) => u.nombre === r.universidad);
      const gastoPct = uni ? Number(uni.gastoAdmin) || 0 : 0;
      const gastoAdmin = montoBase * (gastoPct / 100);
      const totalPorPasante = montoBase + iaposBase + gastoAdmin;
      const cantidad = Number(r.cantidadPasantes) || 0;
      return { carrera: r.carrera, universidad: r.universidad, cantidadPasantes: cantidad, gastoPct, gastoAdmin, totalPorPasante, subtotal: totalPorPasante * cantidad };
    });
    const totalMes = porRenglon.reduce((acc, r) => acc + r.subtotal, 0);
    totalGeneral += totalMes;
    detalleMeses.push({ year, month, estimado: asig.estimado, esProrrateado, diasTrabajados, montoBase, iaposBase, porRenglon, totalMes });
  }

  return { meses: mesesACalcular, mesesHastaFinAno, detalleMeses, totalGeneral, cantidadPasantesTotal, faltanDatos, renglones };
}

function exportPresupuestoExcel(t, presupuesto) {
  const wb = XLSX.utils.book_new();

  // Hoja 1: detalle por mes y por renglón (carrera/universidad)
  const filasDetalle = [];
  presupuesto.detalleMeses.forEach((d) => {
    if (d.sinDato) {
      filasDetalle.push({ Mes: monthLabel(d.year, d.month) + " (sin valor cargado)" });
      return;
    }
    d.porRenglon.forEach((r) => {
      filasDetalle.push({
        Mes: monthLabel(d.year, d.month) + (d.estimado ? " (estimado)" : "") + (d.esProrrateado ? ` — proporcional (${d.diasTrabajados} días)` : ""),
        Carrera: r.carrera,
        "Universidad/Instituto": r.universidad,
        "Cantidad de pasantes": r.cantidadPasantes,
        "Asignación estímulo (c/u)": Number(d.montoBase.toFixed(2)),
        [`IAPOS (${IAPOS_PCT}%, c/u)`]: Number(d.iaposBase.toFixed(2)),
        "% Gasto admin.": r.gastoPct,
        "Gasto admin. (c/u)": Number(r.gastoAdmin.toFixed(2)),
        "Total por pasante": Number(r.totalPorPasante.toFixed(2)),
        "Subtotal renglón": Number(r.subtotal.toFixed(2)),
      });
    });
  });
  filasDetalle.push({ Mes: "TOTAL GENERAL", "Subtotal renglón": Number(presupuesto.totalGeneral.toFixed(2)) });
  const wsDetalle = XLSX.utils.json_to_sheet(filasDetalle);
  XLSX.utils.book_append_sheet(wb, wsDetalle, "Detalle por mes");

  // Hoja 2: resumen por mes
  const filasResumen = presupuesto.detalleMeses.map((d) => ({
    Mes: monthLabel(d.year, d.month),
    "Total del mes": d.sinDato ? "" : Number(d.totalMes.toFixed(2)),
  }));
  filasResumen.push({ Mes: "TOTAL", "Total del mes": Number(presupuesto.totalGeneral.toFixed(2)) });
  const wsResumen = XLSX.utils.json_to_sheet(filasResumen);
  XLSX.utils.book_append_sheet(wb, wsResumen, "Resumen mensual");

  XLSX.writeFile(wb, `presupuesto_${t.expediente || "tramite"}.xlsx`);
}

function PasswordModal({ label, onSubmit, onCancel, error }) {
  const [value, setValue] = useState("");
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,74,0.35)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 90 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: 22, width: 320, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
        <div style={{ fontSize: 13.5, fontWeight: 700, color: "#1B2A4A", marginBottom: 4 }}>Contraseña requerida</div>
        <div style={{ fontSize: 12.5, color: "#5B6158", marginBottom: 12 }}>Para {label}.</div>
        <input
          type="password"
          autoFocus
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && onSubmit(value)}
          style={{ width: "100%", padding: "9px 11px", borderRadius: 6, border: "1px solid " + (error ? "#A6432D" : "#DADCD6"), fontSize: 14, boxSizing: "border-box", marginBottom: error ? 6 : 14 }}
        />
        {error && <div style={{ fontSize: 12, color: "#A6432D", marginBottom: 12 }}>{error}</div>}
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
          <button onClick={onCancel} style={{ padding: "8px 14px", borderRadius: 6, border: "1px solid #DADCD6", background: "#fff", fontSize: 13.5, cursor: "pointer" }}>
            Cancelar
          </button>
          <button onClick={() => onSubmit(value)} style={{ padding: "8px 14px", borderRadius: 6, border: "none", background: "#1B2A4A", color: "#fff", fontSize: 13.5, cursor: "pointer", fontWeight: 600 }}>
            Ingresar
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------- App principal ----------
const DEFAULT_CATALOGOS = {
  jurisdicciones: ["13 - MGIP", "06 - Poder Ejecutivo"],
  habilitaciones: [],
  universidades: [], // [{ nombre, gastoAdmin: number, centralizada: bool }]
  asignaciones: [], // [{ id, year, month, monto, norma, observacion, creadoEl }]
  centralizadaNombre: "",
  carreras: [],
  estadosTramite: [],
  seguridad: { pwdEdicion: "", pwdLiquidacion: "", pwdCatalogo: "" },
};

const IAPOS_PCT = 6;

export default function App() {
  const [pasantias, setPasantias] = useState([]);
  const [catalogos, setCatalogos] = useState(DEFAULT_CATALOGOS);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("resumen");
  const [modal, setModal] = useState(null); // {mode:'new'|'edit', data}
  const [dniError, setDniError] = useState("");
  const [search, setSearch] = useState("");
  const [filterEstado, setFilterEstado] = useState("");
  const [now, setNow] = useState(new Date());
  const [novMes, setNovMes] = useState(new Date().getMonth());
  const [novYear, setNovYear] = useState(new Date().getFullYear());
  const [repDims, setRepDims] = useState({ jurisdiccion: true, habilitacionPagadora: false, universidad: false, carrera: false });
  const [repPeriodo, setRepPeriodo] = useState(false);
  const [repMes, setRepMes] = useState(new Date().getMonth());
  const [repYear, setRepYear] = useState(new Date().getFullYear());
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [printData, setPrintData] = useState(null);
  const [newCatItem, setNewCatItem] = useState({ jurisdicciones: "", habilitaciones: "", carreras: "", estadosTramite: "" });
  const [tramites, setTramites] = useState([]);
  const [nuevaUni, setNuevaUni] = useState({ nombre: "", gastoAdmin: "0", centralizada: false });
  const [nuevaAsig, setNuevaAsig] = useState({ year: new Date().getFullYear(), month: new Date().getMonth(), monto: "", norma: "", observacion: "" });
  const [tramiteModal, setTramiteModal] = useState(null);
  const [confirmDeleteTramite, setConfirmDeleteTramite] = useState(null);
  const [liqSub, setLiqSub] = useState("mensual");
  const [liqMes, setLiqMes] = useState(new Date().getMonth());
  const [liqYear, setLiqYear] = useState(new Date().getFullYear());
  const [compMes, setCompMes] = useState(new Date().getMonth());
  const [compYear, setCompYear] = useState(new Date().getFullYear());
  const [compNuevoValor, setCompNuevoValor] = useState("");
  const [compCalculada, setCompCalculada] = useState(false);
  const [eeMap, setEeMap] = useState({});
  const [liquidaciones, setLiquidaciones] = useState([]);
  const [liqPrint, setLiqPrint] = useState(null);
  const [historialAbierto, setHistorialAbierto] = useState(null);
  const [presupuestoModal, setPresupuestoModal] = useState(null);
  const [notificaciones, setNotificaciones] = useState({});
  const [pwdVisible, setPwdVisible] = useState({ pwdEdicion: false, pwdLiquidacion: false, pwdCatalogo: false });
  const [tramiteSearch, setTramiteSearch] = useState("");

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("pasantias-v1", true);
        if (res?.value) setPasantias(JSON.parse(res.value));
      } catch (e) {}
      try {
        const resCat = await window.storage.get("catalogos-v1", true);
        if (resCat?.value) setCatalogos(normalizeCatalogos(JSON.parse(resCat.value)));
        else await window.storage.set("catalogos-v1", JSON.stringify(DEFAULT_CATALOGOS), true);
      } catch (e) {}
      try {
        const resTram = await window.storage.get("tramites-v1", true);
        if (resTram?.value) setTramites(JSON.parse(resTram.value));
      } catch (e) {}
      try {
        const resEe = await window.storage.get("ee-v1", true);
        if (resEe?.value) setEeMap(JSON.parse(resEe.value));
      } catch (e) {}
      try {
        const resLiq = await window.storage.get("liquidaciones-v1", true);
        if (resLiq?.value) setLiquidaciones(JSON.parse(resLiq.value));
      } catch (e) {}
      try {
        const resNotif = await window.storage.get("notificaciones-v1", true);
        if (resNotif?.value) setNotificaciones(JSON.parse(resNotif.value));
      } catch (e) {}
      setLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set("pasantias-v1", JSON.stringify(pasantias), true).catch(() => {});
  }, [pasantias, loaded]);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set("catalogos-v1", JSON.stringify(catalogos), true).catch(() => {});
  }, [catalogos, loaded]);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set("tramites-v1", JSON.stringify(tramites), true).catch(() => {});
  }, [tramites, loaded]);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set("ee-v1", JSON.stringify(eeMap), true).catch(() => {});
  }, [eeMap, loaded]);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set("liquidaciones-v1", JSON.stringify(liquidaciones), true).catch(() => {});
  }, [liquidaciones, loaded]);

  useEffect(() => {
    if (!loaded) return;
    window.storage.set("notificaciones-v1", JSON.stringify(notificaciones), true).catch(() => {});
  }, [notificaciones, loaded]);

  const addCatalogo = (key, value) => {
    setCatalogos((c) => (c[key].includes(value) ? c : { ...c, [key]: [...c[key], value] }));
  };
  const removeCatalogo = (key, value) => {
    setCatalogos((c) => ({ ...c, [key]: c[key].filter((v) => v !== value) }));
  };

  const addUniversidad = () => {
    if (!nuevaUni.nombre.trim()) return;
    setCatalogos((c) => ({
      ...c,
      universidades: [...c.universidades, { nombre: nuevaUni.nombre.trim(), gastoAdmin: Number(nuevaUni.gastoAdmin) || 0, centralizada: nuevaUni.centralizada }],
    }));
    setNuevaUni({ nombre: "", gastoAdmin: "0", centralizada: false });
  };
  const removeUniversidad = (nombre) => setCatalogos((c) => ({ ...c, universidades: c.universidades.filter((u) => u.nombre !== nombre) }));
  const updateUniversidad = (nombre, patch) =>
    setCatalogos((c) => ({ ...c, universidades: c.universidades.map((u) => (u.nombre === nombre ? { ...u, ...patch } : u)) }));

  const setCentralizadaNombre = (value) => setCatalogos((c) => ({ ...c, centralizadaNombre: value }));

  const addAsignacion = () => {
    const monto = Number(nuevaAsig.monto);
    if (!monto || monto <= 0) return;
    const nueva = {
      id: uid(),
      year: Number(nuevaAsig.year),
      month: Number(nuevaAsig.month),
      monto,
      norma: nuevaAsig.norma || "",
      observacion: nuevaAsig.observacion || "",
      creadoEl: new Date().toISOString(),
    };
    setCatalogos((c) => ({ ...c, asignaciones: [...c.asignaciones, nueva] }));
    setNuevaAsig((n) => ({ ...n, monto: "", norma: "", observacion: "" }));
  };
  const removeAsignacionId = (id) => setCatalogos((c) => ({ ...c, asignaciones: c.asignaciones.filter((a) => a.id !== id) }));

  const [restoreMsg, setRestoreMsg] = useState("");

  const descargarBackup = () => {
    const payload = { version: 3, exportadoEl: new Date().toISOString(), pasantias, catalogos, tramites, eeMap, liquidaciones, notificaciones };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `backup_pasantias_${todayStr()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const [confirmAction, setConfirmAction] = useState(null); // {message, onConfirm}
  const [desbloqueado, setDesbloqueado] = useState({ edicion: false, liquidacion: false, catalogo: false });
  const [passwordPrompt, setPasswordPrompt] = useState(null); // {categoria, accion}
  const [passwordError, setPasswordError] = useState("");

  const CATEGORIA_LABEL = { edicion: "carga y edición de pasantías", liquidacion: "confirmar liquidaciones", catalogo: "modificar catálogos" };
  const CATEGORIA_PWD_KEY = { edicion: "pwdEdicion", liquidacion: "pwdLiquidacion", catalogo: "pwdCatalogo" };

  const ejecutarProtegido = (categoria, accion) => {
    const pwd = catalogos.seguridad?.[CATEGORIA_PWD_KEY[categoria]];
    if (!pwd || desbloqueado[categoria]) {
      accion();
      return;
    }
    setPasswordPrompt({ categoria, accion });
  };

  const restaurarBackup = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = JSON.parse(evt.target.result);
        const okPasantias = Array.isArray(data.pasantias) ? data.pasantias.length : 0;
        setConfirmAction({
          message: `Se va a REEMPLAZAR todo lo que hay cargado ahora por el contenido de este backup (${okPasantias} pasantías). ¿Confirmás?`,
          onConfirm: () => {
            // Compatible con backups viejos (version 1 y 2) que no tenían eeMap/liquidaciones/notificaciones.
            if (Array.isArray(data.pasantias)) setPasantias(data.pasantias);
            if (data.catalogos) setCatalogos(normalizeCatalogos(data.catalogos));
            if (Array.isArray(data.tramites)) setTramites(data.tramites);
            if (data.eeMap && typeof data.eeMap === "object") setEeMap(data.eeMap);
            if (Array.isArray(data.liquidaciones)) setLiquidaciones(data.liquidaciones);
            if (data.notificaciones && typeof data.notificaciones === "object") setNotificaciones(data.notificaciones);
            setRestoreMsg(`Restaurado: ${okPasantias} pasantías.`);
            setTimeout(() => setRestoreMsg(""), 4000);
          },
        });
      } catch (err) {
        setRestoreMsg("Error: el archivo no es un backup válido.");
        setTimeout(() => setRestoreMsg(""), 5000);
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };
  const saveTramite = (t) => {
    setTramites((list) => {
      const exists = list.some((x) => x.id === t.id);
      return exists ? list.map((x) => (x.id === t.id ? t : x)) : [...list, t];
    });
    setTramiteModal(null);
  };
  const deleteTramite = (id) => {
    setTramites((list) => list.filter((x) => x.id !== id));
    setConfirmDeleteTramite(null);
  };

  const savePasantia = (p) => {
    const originalId = modal?.data?.id || null;
    const duplicate = pasantias.some((x) => x.dni === p.dni && x.id !== originalId);
    if (duplicate) {
      setDniError("Ya existe una pasantía cargada con este DNI.");
      return;
    }
    setDniError("");
    setPasantias((list) => {
      let base = list;
      if (originalId && originalId !== p.id) base = base.filter((x) => x.id !== originalId);
      const exists = base.some((x) => x.id === p.id);
      return exists ? base.map((x) => (x.id === p.id ? p : x)) : [...base, p];
    });
    setModal(null);
  };
  const deletePasantia = (id) => {
    setPasantias((list) => list.filter((x) => x.id !== id));
    setConfirmDelete(null);
  };

  const enriched = useMemo(
    () => pasantias.map((p) => ({ ...p, _estado: getEstado(p, new Date()), _vencEfectiva: fechaVencimientoEfectiva(p) })),
    [pasantias, now]
  );

  const alertas = useMemo(
    () =>
      enriched
        .filter((p) => p._estado === "por_vencer_renovacion" || p._estado === "por_vencer_final")
        .sort((a, b) => a._vencEfectiva - b._vencEfectiva),
    [enriched]
  );

  const filtered = useMemo(() => {
    return enriched.filter((p) => {
      const matchSearch =
        !search ||
        [p.nombre, p.lugarTrabajo, p.universidad, p.jurisdiccion, p.habilitacionPagadora].join(" ").toLowerCase().includes(search.toLowerCase());
      const matchEstado = !filterEstado || p._estado === filterEstado;
      return matchSearch && matchEstado;
    });
  }, [enriched, search, filterEstado]);

  const conteos = useMemo(() => {
    const c = { vigente: 0, por_vencer_renovacion: 0, por_vencer_final: 0, suspendida: 0, vencida: 0, renuncia: 0 };
    enriched.forEach((p) => c[p._estado]++);
    return c;
  }, [enriched]);

  const novedadesMes = useMemo(() => {
    const items = [];
    pasantias.forEach((p) => {
      novedadesDe(p).forEach((ev) => {
        const d = toDate(ev.fecha);
        if (d.getFullYear() === novYear && d.getMonth() === novMes) {
          items.push({ ...ev, pasante: p.nombre, universidad: p.universidad, jurisdiccion: p.jurisdiccion });
        }
      });
    });
    return items.sort((a, b) => toDate(a.fecha) - toDate(b.fecha));
  }, [pasantias, novMes, novYear]);

  const novedadesConteo = useMemo(() => {
    const c = { alta: 0, renovacion: 0, suspension_inicio: 0, renuncia: 0, baja_vencimiento: 0 };
    novedadesMes.forEach((ev) => {
      if (c[ev.tipo] !== undefined) c[ev.tipo]++;
    });
    return c;
  }, [novedadesMes]);

  const reportBase = useMemo(() => {
    if (!repPeriodo) return enriched;
    return enriched.filter((p) => vigenteEnMes(p, repYear, repMes));
  }, [enriched, repPeriodo, repMes, repYear]);

  const reportGroups = useMemo(() => {
    const dims = Object.keys(repDims).filter((k) => repDims[k]);
    if (dims.length === 0) return null;
    const map = new Map();
    reportBase.forEach((p) => {
      const key = dims.map((d) => p[d] || "(sin dato)").join(" · ");
      if (!map.has(key)) map.set(key, { key, count: 0, items: [] });
      const g = map.get(key);
      g.count++;
      g.items.push(p);
    });
    return Array.from(map.values()).sort((a, b) => b.count - a.count);
  }, [reportBase, repDims]);

  const maxCount = reportGroups ? Math.max(...reportGroups.map((g) => g.count), 1) : 1;

  const valorAsignacionMes = useMemo(() => getAsignacion(catalogos, liqYear, liqMes), [catalogos, liqYear, liqMes]);
  const liquidacionMensual = useMemo(() => {
    if (!valorAsignacionMes) return null;
    return calcularLiquidacion(pasantias, catalogos, liqYear, liqMes, valorAsignacionMes);
  }, [pasantias, catalogos, liqYear, liqMes, valorAsignacionMes]);
  const baseMensual = useMemo(() => liquidacionConfirmadaAcumulada(liquidaciones, liqYear, liqMes), [liquidaciones, liqYear, liqMes]);
  const diffMensual = useMemo(() => {
    if (!liquidacionMensual || !baseMensual.tieneRegistros) return [];
    return calcularDiferenciaLiquidacion(liquidacionMensual, baseMensual.filas);
  }, [liquidacionMensual, baseMensual]);

  const baseComp = useMemo(() => liquidacionConfirmadaAcumulada(liquidaciones, compYear, compMes), [liquidaciones, compYear, compMes]);
  const valorCatalogoComp = useMemo(() => getAsignacion(catalogos, compYear, compMes), [catalogos, compYear, compMes]);
  const valorUsadoComp = compNuevoValor !== "" ? Number(compNuevoValor) : valorCatalogoComp;
  const liquidacionActualComp = useMemo(() => {
    if (!valorUsadoComp) return null;
    return calcularLiquidacion(pasantias, catalogos, compYear, compMes, valorUsadoComp);
  }, [pasantias, catalogos, compYear, compMes, valorUsadoComp]);
  const liquidacionComp = useMemo(() => {
    if (!compCalculada || !liquidacionActualComp || !baseComp.tieneRegistros) return null;
    const filasDelta = calcularDiferenciaLiquidacion(liquidacionActualComp, baseComp.filas);
    return agruparFilas(filasDelta);
  }, [compCalculada, liquidacionActualComp, baseComp]);

  const fontImport = `
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Serif:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');
    @media print {
      body * { visibility: hidden; }
      .print-area, .print-area * { visibility: visible; }
      .print-area { position: fixed !important; inset: 0 !important; }
      .no-print { display: none !important; }
    }
  `;

  const navItems = [
    { id: "resumen", label: "Resumen" },
    { id: "listado", label: "Pasantías" },
    { id: "tramites", label: "Trámites" },
    { id: "novedades", label: "Novedades" },
    { id: "reportes", label: "Reportes" },
    { id: "liquidacion", label: "Liquidación" },
    { id: "catalogos", label: "Catálogos" },
  ];

  const EstadoPill = ({ estado }) => {
    const e = ESTADOS[estado];
    return (
      <span style={{ display: "inline-flex", alignItems: "center", gap: 5, padding: "3px 9px", borderRadius: 20, fontSize: 12, fontWeight: 600, color: e.color, background: e.bg }}>
        <span style={{ width: 6, height: 6, borderRadius: "50%", background: e.color }} />
        {e.label}
      </span>
    );
  };

  return (
    <div style={{ fontFamily: "'IBM Plex Sans', sans-serif", background: "#F6F7F5", minHeight: "100vh", color: "#20241F" }}>
      <style>{fontImport}</style>

      {/* Header */}
      <div style={{ background: "#1B2A4A", padding: "22px 28px 0" }}>
        <div style={{ maxWidth: 1180, margin: "0 auto" }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <Landmark size={22} color="#D8CBA3" />
              <h1 style={{ fontFamily: "'IBM Plex Serif', serif", fontSize: 21, fontWeight: 600, color: "#fff", margin: 0, letterSpacing: "0.01em" }}>
                Gestión de Pasantías
              </h1>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              {restoreMsg && <span style={{ fontSize: 12, color: "#D8CBA3" }}>{restoreMsg}</span>}
              <button
                onClick={descargarBackup}
                title="Descargar copia de seguridad completa (JSON)"
                style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 12px", background: "rgba(255,255,255,0.1)", color: "#fff", border: "1px solid rgba(255,255,255,0.25)", borderRadius: 6, fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}
              >
                <Save size={13} /> Copia de seguridad
              </button>
              <label
                htmlFor="restore-backup-input"
                title="Restaurar desde una copia de seguridad"
                style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 12px", background: "rgba(255,255,255,0.1)", color: "#fff", border: "1px solid rgba(255,255,255,0.25)", borderRadius: 6, fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}
              >
                <UploadCloud size={13} /> Restaurar
              </label>
              <input id="restore-backup-input" type="file" accept="application/json,.json" onChange={restaurarBackup} style={{ display: "none" }} />
            </div>
          </div>
          <div style={{ display: "flex", gap: 4 }}>
            {navItems.map((n) => (
              <button
                key={n.id}
                onClick={() => setTab(n.id)}
                style={{
                  padding: "10px 18px",
                  background: tab === n.id ? "#F6F7F5" : "transparent",
                  color: tab === n.id ? "#1B2A4A" : "#C7CDDB",
                  border: "none",
                  borderRadius: "8px 8px 0 0",
                  fontSize: 14,
                  fontWeight: 600,
                  cursor: "pointer",
                  fontFamily: "'IBM Plex Sans', sans-serif",
                }}
              >
                {n.label}
              </button>
            ))}
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1180, margin: "0 auto", padding: "26px 28px 60px" }}>
        {/* -------- RESUMEN -------- */}
        {tab === "resumen" && (
          <div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 12, marginBottom: 26 }}>
              {[
                ["vigente", "Vigentes"],
                ["por_vencer_renovacion", "Por vencer (renovación)"],
                ["por_vencer_final", "Por vencer (final)"],
                ["suspendida", "Suspendidas"],
                ["vencida", "Vencidas"],
              ].map(([k, label]) => (
                <div key={k} style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, padding: 16 }}>
                  <div style={{ fontSize: 27, fontWeight: 700, color: ESTADOS[k].color, fontFamily: "'IBM Plex Serif', serif" }}>{conteos[k]}</div>
                  <div style={{ fontSize: 12.5, color: "#5B6158", marginTop: 2 }}>{label}</div>
                </div>
              ))}
            </div>

            <div style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, overflow: "hidden" }}>
              <div style={{ padding: "14px 18px", borderBottom: "1px solid #E7E9E4", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <AlertTriangle size={17} color="#B8862F" />
                  <span style={{ fontFamily: "'IBM Plex Serif', serif", fontWeight: 600, fontSize: 15 }}>Alertas de vencimiento (próximos 45 días)</span>
                </div>
                {alertas.length > 0 && (
                  <button
                    onClick={() =>
                      setPrintData(alertas.map((p) => ({ p, estado: p._estado, restantes: daysBetween(new Date(new Date().setHours(0, 0, 0, 0)), p._vencEfectiva) })))
                    }
                    style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 12px", background: "#1B2A4A", color: "#fff", border: "none", borderRadius: 6, fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}
                  >
                    <Printer size={13} /> Generar aviso general (PDF)
                  </button>
                )}
              </div>
              {alertas.length === 0 && (
                <div style={{ padding: 24, textAlign: "center", color: "#8A9088", fontSize: 14 }}>No hay pasantías próximas a vencer.</div>
              )}
              {alertas.map((p) => {
                const restantes = daysBetween(new Date(new Date().setHours(0, 0, 0, 0)), p._vencEfectiva);
                const notifKey = `${p.id}-${p._vencEfectiva.toISOString().slice(0, 10)}`;
                const notificada = !!notificaciones[notifKey];
                return (
                  <div key={p.id} style={{ padding: "12px 18px", borderBottom: "1px solid #F0F1EE", display: "flex", justifyContent: "space-between", alignItems: "center", opacity: notificada ? 0.55 : 1 }}>
                    <div onClick={() => ejecutarProtegido("edicion", () => setModal({ mode: "edit", data: p }))} style={{ cursor: "pointer" }}>
                      <div style={{ fontWeight: 600, fontSize: 14 }}>{p.nombre}</div>
                      <div style={{ fontSize: 12.5, color: "#8A9088", marginTop: 2 }}>
                        {p.universidad} · {p.jurisdiccion} ·{" "}
                        {p._estado === "por_vencer_renovacion" ? "iniciar gestión de renovación" : "finalización definitiva"}
                      </div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                      <div style={{ textAlign: "right" }}>
                        <div style={{ fontSize: 13, fontWeight: 700, color: "#B8862F" }}>{restantes} días</div>
                        <div style={{ fontSize: 11.5, color: "#8A9088" }}>vence {fmt(p._vencEfectiva.toISOString().slice(0, 10))}</div>
                      </div>
                      <button
                        onClick={() => setPrintData([{ p, estado: p._estado, restantes }])}
                        title="Generar aviso individual (PDF)"
                        style={{ background: "none", border: "1px solid #DADCD6", borderRadius: 6, cursor: "pointer", color: "#1B2A4A", padding: 7 }}
                      >
                        <Printer size={14} />
                      </button>
                      <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11.5, color: "#5B6158", cursor: "pointer" }}>
                        <input
                          type="checkbox"
                          checked={notificada}
                          onChange={(e) => setNotificaciones((n) => ({ ...n, [notifKey]: e.target.checked }))}
                        />
                        Notificada
                      </label>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* -------- LISTADO -------- */}
        {tab === "listado" && (
          <div>
            <div style={{ display: "flex", gap: 10, marginBottom: 16, alignItems: "center" }}>
              <div style={{ position: "relative", flex: 1 }}>
                <Search size={15} color="#8A9088" style={{ position: "absolute", left: 12, top: 11 }} />
                <input
                  placeholder="Buscar por nombre, lugar, universidad, jurisdicción…"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  style={{ width: "100%", padding: "9px 12px 9px 34px", borderRadius: 7, border: "1px solid #DADCD6", fontSize: 14, boxSizing: "border-box" }}
                />
              </div>
              <select value={filterEstado} onChange={(e) => setFilterEstado(e.target.value)} style={{ padding: "9px 10px", borderRadius: 7, border: "1px solid #DADCD6", fontSize: 13.5 }}>
                <option value="">Todos los estados</option>
                {Object.entries(ESTADOS).map(([k, v]) => (
                  <option key={k} value={k}>{v.label}</option>
                ))}
              </select>
              <button
                onClick={() =>
                  exportExcel({
                    detalle: filtered,
                    nombreArchivo: `pasantias_${todayStr()}.xlsx`,
                    hojaDetalle: "Pasantías",
                  })
                }
                style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 14px", background: "#fff", color: "#1B2A4A", border: "1px solid #DADCD6", borderRadius: 7, fontSize: 14, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" }}
              >
                <Download size={15} /> Excel
              </button>
              <button
                onClick={() => ejecutarProtegido("edicion", () => setModal({ mode: "new", data: null }))}
                style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: "#1B2A4A", color: "#fff", border: "none", borderRadius: 7, fontSize: 14, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" }}
              >
                <Plus size={16} /> Nueva pasantía
              </button>
            </div>

            {filtered.length === 0 && (
              <div style={{ background: "#fff", border: "1px dashed #DADCD6", borderRadius: 10, padding: 40, textAlign: "center", color: "#8A9088" }}>
                {pasantias.length === 0 ? "Todavía no cargaste ninguna pasantía. Empezá con \"Nueva pasantía\"." : "No hay resultados para ese filtro."}
              </div>
            )}

            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {filtered.map((p) => (
                <div key={p.id} style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, padding: "14px 16px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
                    <div>
                      <div style={{ fontWeight: 700, fontSize: 15, fontFamily: "'IBM Plex Serif', serif" }}>{p.nombre}</div>
                      <div style={{ display: "flex", gap: 14, marginTop: 4, fontSize: 12.5, color: "#5B6158", flexWrap: "wrap" }}>
                        {p.lugarTrabajo && <span style={{ display: "flex", alignItems: "center", gap: 4 }}><MapPin size={12} />{p.lugarTrabajo}</span>}
                        {p.universidad && <span style={{ display: "flex", alignItems: "center", gap: 4 }}><GraduationCap size={12} />{p.universidad}{p.carrera ? ` · ${p.carrera}` : ""}</span>}
                        {p.jurisdiccion && <span style={{ display: "flex", alignItems: "center", gap: 4 }}><Building2 size={12} />{p.jurisdiccion}</span>}
                        {p.expediente && <span style={{ display: "flex", alignItems: "center", gap: 4 }}><FileText size={12} />{p.expediente}</span>}
                      </div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <EstadoPill estado={p._estado} />
                      <button onClick={() => ejecutarProtegido("edicion", () => setModal({ mode: "edit", data: p }))} style={{ background: "none", border: "none", cursor: "pointer", color: "#5B6158", padding: 4 }}>
                        <Edit2 size={15} />
                      </button>
                      <button onClick={() => setConfirmDelete(p.id)} style={{ background: "none", border: "none", cursor: "pointer", color: "#A6432D", padding: 4 }}>
                        <Trash2 size={15} />
                      </button>
                    </div>
                  </div>
                  <Timeline p={p} />
                  <div style={{ display: "flex", justifyContent: "space-between", marginTop: 5, fontSize: 11, color: "#8A9088" }}>
                    <span>{fmt(p.periodo1.inicio)}</span>
                    <span>{fmt(p.renuncia?.fecha || p._vencEfectiva.toISOString().slice(0, 10))}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* -------- NOVEDADES -------- */}
        {tab === "novedades" && (
          <div>
            <div style={{ display: "flex", gap: 10, marginBottom: 18, alignItems: "center" }}>
              <select value={novMes} onChange={(e) => setNovMes(Number(e.target.value))} style={{ padding: "9px 10px", borderRadius: 7, border: "1px solid #DADCD6", fontSize: 13.5 }}>
                {Array.from({ length: 12 }).map((_, i) => (
                  <option key={i} value={i}>{new Date(2000, i, 1).toLocaleDateString("es-AR", { month: "long" })}</option>
                ))}
              </select>
              <select value={novYear} onChange={(e) => setNovYear(Number(e.target.value))} style={{ padding: "9px 10px", borderRadius: 7, border: "1px solid #DADCD6", fontSize: 13.5 }}>
                {Array.from({ length: 7 }).map((_, i) => {
                  const y = new Date().getFullYear() - 2 + i;
                  return <option key={y} value={y}>{y}</option>;
                })}
              </select>
              <div style={{ fontFamily: "'IBM Plex Serif', serif", fontWeight: 600, fontSize: 15, color: "#1B2A4A", textTransform: "capitalize" }}>
                {monthLabel(novYear, novMes)} — {novedadesMes.length} novedad{novedadesMes.length !== 1 ? "es" : ""}
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 12, marginBottom: 22 }}>
              {[
                ["alta", "Altas"],
                ["renovacion", "Renovaciones"],
                ["suspension_inicio", "Suspensiones"],
                ["renuncia", "Renuncias"],
                ["baja_vencimiento", "Bajas por vencimiento"],
              ].map(([k, label]) => {
                const t = TIPO_LABEL[k];
                return (
                  <div key={k} style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, padding: 16 }}>
                    <div style={{ fontSize: 27, fontWeight: 700, color: t.color, fontFamily: "'IBM Plex Serif', serif" }}>{novedadesConteo[k]}</div>
                    <div style={{ fontSize: 12.5, color: "#5B6158", marginTop: 2 }}>{label}</div>
                  </div>
                );
              })}
            </div>

            {novedadesMes.length === 0 && (
              <div style={{ background: "#fff", border: "1px dashed #DADCD6", borderRadius: 10, padding: 40, textAlign: "center", color: "#8A9088" }}>
                No hay novedades registradas para este mes.
              </div>
            )}

            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {novedadesMes.map((ev, i) => {
                const t = TIPO_LABEL[ev.tipo];
                const Icon = t.icon;
                return (
                  <div key={i} style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 9, padding: "12px 16px", display: "flex", alignItems: "center", gap: 14 }}>
                    <div style={{ width: 34, height: 34, borderRadius: 8, background: t.color + "1A", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                      <Icon size={16} color={t.color} />
                    </div>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 13.5, fontWeight: 600 }}>{ev.pasante}</div>
                      <div style={{ fontSize: 12.5, color: "#5B6158", marginTop: 1 }}>{ev.detalle}</div>
                      <div style={{ fontSize: 11.5, color: "#8A9088", marginTop: 1 }}>{ev.universidad} · {ev.jurisdiccion}</div>
                    </div>
                    <div style={{ fontSize: 12.5, fontWeight: 600, color: t.color }}>{fmt(ev.fecha)}</div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* -------- REPORTES -------- */}
        {tab === "reportes" && (
          <div>
            <div style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, padding: 16, marginBottom: 18 }}>
              <div style={{ fontFamily: "'IBM Plex Serif', serif", fontWeight: 600, fontSize: 14.5, marginBottom: 12 }}>Agrupar por</div>
              <div style={{ display: "flex", gap: 18, marginBottom: 16, flexWrap: "wrap" }}>
                {[
                  ["jurisdiccion", "Jurisdicción"],
                  ["habilitacionPagadora", "Habilitación pagadora"],
                  ["universidad", "Universidad / Instituto"],
                  ["carrera", "Carrera"],
                ].map(([k, label]) => (
                  <label key={k} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13.5, cursor: "pointer" }}>
                    <input type="checkbox" checked={repDims[k]} onChange={(e) => setRepDims((d) => ({ ...d, [k]: e.target.checked }))} />
                    {label}
                  </label>
                ))}
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 12, paddingTop: 12, borderTop: "1px solid #F0F1EE" }}>
                <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 13.5, cursor: "pointer", fontWeight: 600 }}>
                  <input type="checkbox" checked={repPeriodo} onChange={(e) => setRepPeriodo(e.target.checked)} />
                  Filtrar por vigencia en un mes específico
                </label>
                {repPeriodo && (
                  <>
                    <select value={repMes} onChange={(e) => setRepMes(Number(e.target.value))} style={{ padding: "7px 9px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13 }}>
                      {Array.from({ length: 12 }).map((_, i) => (
                        <option key={i} value={i}>{new Date(2000, i, 1).toLocaleDateString("es-AR", { month: "long" })}</option>
                      ))}
                    </select>
                    <select value={repYear} onChange={(e) => setRepYear(Number(e.target.value))} style={{ padding: "7px 9px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13 }}>
                      {Array.from({ length: 7 }).map((_, i) => {
                        const y = new Date().getFullYear() - 2 + i;
                        return <option key={y} value={y}>{y}</option>;
                      })}
                    </select>
                  </>
                )}
              </div>
            </div>

            <div style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, padding: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <div style={{ fontSize: 12.5, color: "#8A9088" }}>
                  {reportBase.length} pasantía{reportBase.length !== 1 ? "s" : ""} en el criterio seleccionado
                  {repPeriodo ? ` · vigentes en ${monthLabel(repYear, repMes)}` : ""}
                </div>
                <button
                  onClick={() =>
                    exportExcel({
                      detalle: reportBase,
                      grupos: reportGroups,
                      nombreArchivo: `reporte_pasantias_${todayStr()}.xlsx`,
                      hojaDetalle: "Detalle",
                    })
                  }
                  disabled={reportBase.length === 0}
                  style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: reportBase.length ? "#1B2A4A" : "#C7CDDB", color: "#fff", border: "none", borderRadius: 7, fontSize: 13.5, fontWeight: 600, cursor: reportBase.length ? "pointer" : "not-allowed" }}
                >
                  <Download size={14} /> Descargar Excel
                </button>
              </div>
              {!reportGroups && <div style={{ textAlign: "center", color: "#8A9088", padding: 20 }}>Seleccioná al menos un criterio de agrupación.</div>}
              {reportGroups && reportGroups.length === 0 && <div style={{ textAlign: "center", color: "#8A9088", padding: 20 }}>Sin resultados.</div>}
              {reportGroups && reportGroups.map((g) => (
                <div key={g.key} style={{ marginBottom: 12 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5, marginBottom: 4 }}>
                    <span style={{ fontWeight: 600 }}>{g.key}</span>
                    <span style={{ color: "#1B2A4A", fontWeight: 700 }}>{g.count}</span>
                  </div>
                  <div style={{ height: 8, background: "#EFF0EC", borderRadius: 4, overflow: "hidden" }}>
                    <div style={{ width: `${(g.count / maxCount) * 100}%`, height: "100%", background: "#2F6E5E" }} />
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* -------- TRÁMITES -------- */}
        {tab === "tramites" && (() => {
          const tramitesFiltrados = tramites.filter((t) => {
            if (!tramiteSearch) return true;
            const q = tramiteSearch.toLowerCase();
            const camposRenglones = (t.detalles || []).flatMap((d) => [d.universidad, d.carrera]);
            return [t.expediente, t.lugarTrabajo, t.jurisdiccion, ...camposRenglones].some((v) => (v || "").toLowerCase().includes(q));
          });
          return (
          <div>
            <div style={{ display: "flex", gap: 10, marginBottom: 16, alignItems: "center" }}>
              <div style={{ position: "relative", flex: 1 }}>
                <Search size={15} color="#8A9088" style={{ position: "absolute", left: 12, top: 11 }} />
                <input
                  placeholder="Buscar por expediente, universidad, carrera, lugar de trabajo o jurisdicción…"
                  value={tramiteSearch}
                  onChange={(e) => setTramiteSearch(e.target.value)}
                  style={{ width: "100%", padding: "9px 12px 9px 34px", borderRadius: 7, border: "1px solid #DADCD6", fontSize: 14, boxSizing: "border-box" }}
                />
              </div>
              <button
                onClick={() => ejecutarProtegido("edicion", () => setTramiteModal({ mode: "new", data: null }))}
                style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 16px", background: "#1B2A4A", color: "#fff", border: "none", borderRadius: 7, fontSize: 14, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" }}
              >
                <Plus size={16} /> Nuevo trámite
              </button>
            </div>

            {tramites.length === 0 && (
              <div style={{ background: "#fff", border: "1px dashed #DADCD6", borderRadius: 10, padding: 40, textAlign: "center", color: "#8A9088" }}>
                No hay contrataciones en trámite cargadas todavía.
              </div>
            )}
            {tramites.length > 0 && tramitesFiltrados.length === 0 && (
              <div style={{ background: "#fff", border: "1px dashed #DADCD6", borderRadius: 10, padding: 40, textAlign: "center", color: "#8A9088" }}>
                No hay resultados para esa búsqueda.
              </div>
            )}

            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {tramitesFiltrados.map((t) => (
                <div key={t.id} style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, padding: "14px 16px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                    <div>
                      <div style={{ fontWeight: 700, fontSize: 15, fontFamily: "'IBM Plex Serif', serif" }}>
                        Expediente {t.expediente} · {(t.detalles || []).reduce((acc, d) => acc + (Number(d.cantidadPasantes) || 0), 0)} pasante{(t.detalles || []).reduce((acc, d) => acc + (Number(d.cantidadPasantes) || 0), 0) !== 1 ? "s" : ""} en total
                      </div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 3, marginTop: 6 }}>
                        {(t.detalles || []).filter((d) => d.universidad || d.carrera).map((d) => (
                          <div key={d.id} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: "#5B6158" }}>
                            <GraduationCap size={12} />
                            <span>{d.cantidadPasantes || "—"} × {d.carrera || "(sin carrera)"} — {d.universidad || "(sin universidad)"}</span>
                          </div>
                        ))}
                      </div>
                      <div style={{ display: "flex", gap: 14, marginTop: 6, fontSize: 12.5, color: "#5B6158", flexWrap: "wrap" }}>
                        {t.lugarTrabajo && <span style={{ display: "flex", alignItems: "center", gap: 4 }}><MapPin size={12} />{t.lugarTrabajo}</span>}
                        {t.habilitacionPagadora && <span>{t.habilitacionPagadora}</span>}
                        {t.jurisdiccion && <span>{t.jurisdiccion}</span>}
                      </div>
                      {t.observaciones && <div style={{ fontSize: 12.5, color: "#8A9088", marginTop: 6 }}>{t.observaciones}</div>}
                      {(t.cantidadMeses || t.fechaProbableInicio) && (
                        <div style={{ fontSize: 12, color: "#5B6158", marginTop: 6 }}>
                          {t.fechaProbableInicio && <>Inicio probable: {fmt(t.fechaProbableInicio)} </>}
                          {t.cantidadMeses && <>· {t.cantidadMeses} mes{Number(t.cantidadMeses) !== 1 ? "es" : ""} de contrato</>}
                        </div>
                      )}
                      {(() => {
                        const presupuesto = calcularPresupuestoTramite(t, catalogos);
                        if (!presupuesto) return null;
                        return (
                          <div style={{ marginTop: 8, display: "flex", alignItems: "center", gap: 10 }}>
                            <span style={{ fontSize: 13, fontWeight: 700, color: "#1B2A4A" }}>
                              Presupuesto estimado {presupuesto.detalleMeses[0]?.year}: {moneyFmt(presupuesto.totalGeneral)}
                            </span>
                            {presupuesto.faltanDatos && (
                              <span style={{ fontSize: 11, color: "#B8862F" }}>(faltan valores de asignación para algún mes)</span>
                            )}
                            <button onClick={() => setPresupuestoModal(t)} style={{ padding: "4px 10px", borderRadius: 6, border: "1px solid #DADCD6", background: "#fff", fontSize: 11.5, cursor: "pointer" }}>
                              Ver detalle
                            </button>
                          </div>
                        );
                      })()}
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      {t.estado && (
                        <span style={{ padding: "3px 10px", borderRadius: 20, fontSize: 12, fontWeight: 600, color: colorPorEstado(t.estado).color, background: colorPorEstado(t.estado).bg }}>{t.estado}</span>
                      )}
                      <button onClick={() => ejecutarProtegido("edicion", () => setTramiteModal({ mode: "edit", data: t }))} style={{ background: "none", border: "none", cursor: "pointer", color: "#5B6158", padding: 4 }}>
                        <Edit2 size={15} />
                      </button>
                      <button onClick={() => setConfirmDeleteTramite(t.id)} style={{ background: "none", border: "none", cursor: "pointer", color: "#A6432D", padding: 4 }}>
                        <Trash2 size={15} />
                      </button>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
          );
        })()}

        {/* -------- LIQUIDACIÓN -------- */}
        {tab === "liquidacion" && (
          <div>
            <div style={{ display: "flex", gap: 4, marginBottom: 18 }}>
              {[
                ["mensual", "Liquidación mensual"],
                ["complementaria", "Complementaria"],
                ["historial", "Historial"],
              ].map(([k, label]) => (
                <button
                  key={k}
                  onClick={() => { setLiqSub(k); setHistorialAbierto(null); }}
                  style={{
                    padding: "9px 16px",
                    background: liqSub === k ? "#1B2A4A" : "#fff",
                    color: liqSub === k ? "#fff" : "#1B2A4A",
                    border: "1px solid " + (liqSub === k ? "#1B2A4A" : "#DADCD6"),
                    borderRadius: 7,
                    fontSize: 13.5,
                    fontWeight: 600,
                    cursor: "pointer",
                  }}
                >
                  {label}
                </button>
              ))}
            </div>

            {liqSub === "mensual" && (() => {
              const periodKey = liqPeriodKey("mensual", liqYear, liqMes);
              const hayDiferencias = baseMensual.tieneRegistros && diffMensual.length > 0;
              return (
                <div>
                  <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 18 }}>
                    <select value={liqMes} onChange={(e) => setLiqMes(Number(e.target.value))} style={{ padding: "9px 10px", borderRadius: 7, border: "1px solid #DADCD6", fontSize: 13.5 }}>
                      {Array.from({ length: 12 }).map((_, i) => (
                        <option key={i} value={i}>{new Date(2000, i, 1).toLocaleDateString("es-AR", { month: "long" })}</option>
                      ))}
                    </select>
                    <select value={liqYear} onChange={(e) => setLiqYear(Number(e.target.value))} style={{ padding: "9px 10px", borderRadius: 7, border: "1px solid #DADCD6", fontSize: 13.5 }}>
                      {Array.from({ length: 7 }).map((_, i) => {
                        const y = new Date().getFullYear() - 2 + i;
                        return <option key={y} value={y}>{y}</option>;
                      })}
                    </select>
                    <div style={{ fontFamily: "'IBM Plex Serif', serif", fontWeight: 600, fontSize: 15, color: "#1B2A4A", textTransform: "capitalize" }}>
                      {monthLabel(liqYear, liqMes)}
                    </div>
                    {valorAsignacionMes && (
                      <div style={{ marginLeft: "auto", fontSize: 12.5, color: "#5B6158" }}>
                        Asignación estímulo vigente: <strong style={{ color: "#1B2A4A" }}>{moneyFmt(valorAsignacionMes)}</strong>
                      </div>
                    )}
                  </div>

                  {!valorAsignacionMes ? (
                    <div style={{ background: "#FBF3E4", border: "1px solid #E9D8AE", borderRadius: 10, padding: 24, textAlign: "center", color: "#8A6420" }}>
                      No hay un valor de asignación estímulo cargado para {monthLabel(liqYear, liqMes)}. Cargalo desde la pestaña <strong>Catálogos</strong> (sección "Asignación estímulo mensual") para poder liquidar este mes.
                    </div>
                  ) : (
                    <div>
                      {hayDiferencias && (
                        <div style={{ background: "#FBF3E4", border: "1px solid #E9D8AE", borderRadius: 10, padding: "14px 16px", marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12 }}>
                          <div style={{ fontSize: 13, color: "#8A6420" }}>
                            <strong>{diffMensual.length}</strong> pasantía{diffMensual.length !== 1 ? "s" : ""} cambiaron respecto a la última liquidación confirmada de este mes (altas, bajas o ajustes cargados después). La liquidación de abajo ya no coincide con lo confirmado.
                          </div>
                          <button
                            onClick={() => {
                              setLiqSub("complementaria");
                              setCompYear(liqYear);
                              setCompMes(liqMes);
                              setCompNuevoValor("");
                              setCompCalculada(false);
                            }}
                            style={{ padding: "8px 14px", borderRadius: 7, border: "none", background: "#B8862F", color: "#fff", fontSize: 12.5, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" }}
                          >
                            Ver complementaria de la diferencia
                          </button>
                        </div>
                      )}
                      {baseMensual.tieneRegistros && !hayDiferencias && (
                        <div style={{ background: "#EAF3EF", border: "1px solid #CFE6DC", borderRadius: 10, padding: "10px 16px", marginBottom: 16, fontSize: 12.5, color: "#2F6E5E" }}>
                          Ya confirmada el {fmt(baseMensual.ultimaFecha?.slice(0, 10))}, sin cambios pendientes respecto a los datos actuales.
                        </div>
                      )}
                      <LiquidacionView
                        liq={liquidacionMensual}
                        etiqueta={monthLabel(liqYear, liqMes)}
                        nombreCentralizada={catalogos.centralizadaNombre}
                        eeMap={eeMap[periodKey] || {}}
                        onAsignarEE={(dnis, valor) =>
                          setEeMap((m) => ({ ...m, [periodKey]: { ...(m[periodKey] || {}), ...Object.fromEntries(dnis.map((d) => [d, valor])) } }))
                        }
                        onExport={() =>
                          exportLiquidacionExcel(
                            liquidacionMensual,
                            `liquidacion_${monthLabel(liqYear, liqMes).replace(/\s|·/g, "_")}.xlsx`,
                            monthLabel(liqYear, liqMes),
                            eeMap[periodKey]
                          )
                        }
                        onPrint={(filtro) => setLiqPrint({ liq: liquidacionMensual, etiqueta: monthLabel(liqYear, liqMes), eeMap: eeMap[periodKey], filtro })}
                        onConfirmar={
                          baseMensual.tieneRegistros
                            ? null
                            : () =>
                                ejecutarProtegido("liquidacion", () =>
                                  setConfirmAction({
                                    message: `¿Confirmar la liquidación de ${monthLabel(liqYear, liqMes)}? Va a quedar guardada en el Historial.`,
                                    onConfirm: () => {
                                      setLiquidaciones((list) => [
                                        ...list,
                                        { id: uid(), tipo: "mensual", year: liqYear, month: liqMes, valor: valorAsignacionMes, fechaConfirmacion: new Date().toISOString(), liq: liquidacionMensual, eeMap: eeMap[periodKey] || {} },
                                      ]);
                                    },
                                  })
                                )
                        }
                        confirmado={baseMensual.tieneRegistros}
                        fechaConfirmacion={baseMensual.ultimaFecha}
                      />
                    </div>
                  )}
                </div>
              );
            })()}

            {liqSub === "complementaria" && (() => {
              const periodKey = liqPeriodKey("complementaria", compYear, compMes);
              return (
                <div>
                  <div style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, padding: 16, marginBottom: 18 }}>
                    <div style={{ fontFamily: "'IBM Plex Serif', serif", fontWeight: 600, fontSize: 14.5, marginBottom: 12 }}>Liquidación complementaria</div>
                    <div style={{ fontSize: 12.5, color: "#8A9088", marginBottom: 14 }}>
                      Compara la liquidación de hoy (con los datos y el valor de asignación actuales) contra la última liquidación confirmada de ese mes, y arma la diferencia: altas nuevas, bajas que ya no corresponden, o ajustes de monto — incluye tanto cambios de valor de asignación como cambios en los datos de las pasantías (altas/bajas cargadas después).
                    </div>
                    <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
                      <div>
                        <label style={{ fontSize: 11.5, fontWeight: 600, color: "#5B6158", marginBottom: 4, display: "block" }}>Mes</label>
                        <select value={compMes} onChange={(e) => { setCompMes(Number(e.target.value)); setCompCalculada(false); }} style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13.5 }}>
                          {Array.from({ length: 12 }).map((_, i) => (
                            <option key={i} value={i}>{new Date(2000, i, 1).toLocaleDateString("es-AR", { month: "long" })}</option>
                          ))}
                        </select>
                      </div>
                      <div>
                        <label style={{ fontSize: 11.5, fontWeight: 600, color: "#5B6158", marginBottom: 4, display: "block" }}>Año</label>
                        <select value={compYear} onChange={(e) => { setCompYear(Number(e.target.value)); setCompCalculada(false); }} style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13.5 }}>
                          {Array.from({ length: 7 }).map((_, i) => {
                            const y = new Date().getFullYear() - 2 + i;
                            return <option key={y} value={y}>{y}</option>;
                          })}
                        </select>
                      </div>
                      <div>
                        <label style={{ fontSize: 11.5, fontWeight: 600, color: "#5B6158", marginBottom: 4, display: "block" }}>Última confirmada</label>
                        <div style={{ padding: "8px 10px", fontSize: 13.5, color: baseComp.tieneRegistros ? "#20241F" : "#A6432D" }}>
                          {baseComp.tieneRegistros ? fmt(baseComp.ultimaFecha?.slice(0, 10)) : "No hay ninguna confirmada"}
                        </div>
                      </div>
                      <div>
                        <label style={{ fontSize: 11.5, fontWeight: 600, color: "#5B6158", marginBottom: 4, display: "block" }}>Valor asignación a usar</label>
                        <input
                          type="number"
                          placeholder={valorCatalogoComp ? `Vigente: ${valorCatalogoComp}` : "Sin valor cargado"}
                          value={compNuevoValor}
                          onChange={(e) => { setCompNuevoValor(e.target.value); setCompCalculada(false); }}
                          style={{ padding: "8px 10px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13.5, width: 160 }}
                        />
                        <div style={{ fontSize: 10.5, color: "#8A9088", marginTop: 3 }}>Vacío = usa el valor vigente del catálogo para este mes</div>
                      </div>
                      <button
                        onClick={() => setCompCalculada(true)}
                        disabled={!baseComp.tieneRegistros || !valorUsadoComp}
                        style={{ padding: "9px 16px", borderRadius: 7, border: "none", background: baseComp.tieneRegistros && valorUsadoComp ? "#1B2A4A" : "#C7CDDB", color: "#fff", fontSize: 13.5, fontWeight: 600, cursor: baseComp.tieneRegistros && valorUsadoComp ? "pointer" : "not-allowed" }}
                      >
                        Calcular
                      </button>
                    </div>
                    {!baseComp.tieneRegistros && (
                      <div style={{ marginTop: 12, fontSize: 12, color: "#A6432D" }}>
                        Todavía no hay ninguna liquidación confirmada para {monthLabel(compYear, compMes)}. Primero liquidá y confirmá ese mes desde "Liquidación mensual".
                      </div>
                    )}
                  </div>

                  {compCalculada && liquidacionComp && liquidacionComp.filas.length > 0 && (
                    <LiquidacionView
                      liq={liquidacionComp}
                      etiqueta={`Complementaria ${monthLabel(compYear, compMes)}`}
                      nombreCentralizada={catalogos.centralizadaNombre}
                      eeMap={eeMap[periodKey] || {}}
                      onAsignarEE={(dnis, valor) =>
                        setEeMap((m) => ({ ...m, [periodKey]: { ...(m[periodKey] || {}), ...Object.fromEntries(dnis.map((d) => [d, valor])) } }))
                      }
                      onExport={() =>
                        exportLiquidacionExcel(
                          liquidacionComp,
                          `complementaria_${monthLabel(compYear, compMes).replace(/\s|·/g, "_")}.xlsx`,
                          `Complementaria ${monthLabel(compYear, compMes)}`,
                          eeMap[periodKey]
                        )
                      }
                      onPrint={(filtro) => setLiqPrint({ liq: liquidacionComp, etiqueta: `Complementaria ${monthLabel(compYear, compMes)}`, eeMap: eeMap[periodKey], filtro })}
                      onConfirmar={() =>
                        ejecutarProtegido("liquidacion", () =>
                          setConfirmAction({
                            message: `¿Confirmar la complementaria de ${monthLabel(compYear, compMes)}? Va a quedar guardada en el Historial y va a pasar a formar parte de lo confirmado de ese mes.`,
                            onConfirm: () => {
                              setLiquidaciones((list) => [
                                ...list,
                                { id: uid(), tipo: "complementaria", year: compYear, month: compMes, valor: valorUsadoComp, fechaConfirmacion: new Date().toISOString(), liq: liquidacionComp, eeMap: eeMap[periodKey] || {} },
                              ]);
                            },
                          })
                        )
                      }
                    />
                  )}
                  {compCalculada && liquidacionComp && liquidacionComp.filas.length === 0 && (
                    <div style={{ background: "#fff", border: "1px dashed #DADCD6", borderRadius: 10, padding: 24, textAlign: "center", color: "#8A9088" }}>
                      No hay diferencias entre lo confirmado y la situación actual para este mes.
                    </div>
                  )}
                </div>
              );
            })()}

            {liqSub === "historial" && (
              <div>
                {liquidaciones.length === 0 && (
                  <div style={{ background: "#fff", border: "1px dashed #DADCD6", borderRadius: 10, padding: 30, textAlign: "center", color: "#8A9088" }}>
                    Todavía no confirmaste ninguna liquidación.
                  </div>
                )}
                {!historialAbierto && (
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    {liquidaciones
                      .slice()
                      .sort((a, b) => b.fechaConfirmacion.localeCompare(a.fechaConfirmacion))
                      .map((r) => (
                        <div key={r.id} style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, padding: "13px 16px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                          <div>
                            <div style={{ fontWeight: 700, fontSize: 14, textTransform: "capitalize" }}>
                              {r.tipo === "mensual" ? "Mensual" : "Complementaria"} — {monthLabel(r.year, r.month)}
                            </div>
                            <div style={{ fontSize: 12, color: "#8A9088", marginTop: 2 }}>
                              Confirmada el {fmt(r.fechaConfirmacion.slice(0, 10))} · Valor: {moneyFmt(r.valor)} · {r.liq.filas.length} pasantías
                            </div>
                          </div>
                          <div style={{ display: "flex", gap: 8 }}>
                            <button onClick={() => setHistorialAbierto(r)} style={{ padding: "7px 12px", borderRadius: 6, border: "1px solid #DADCD6", background: "#fff", fontSize: 12.5, cursor: "pointer" }}>Ver</button>
                            <button
                              onClick={() =>
                                setConfirmAction({
                                  message: "¿Eliminar esta liquidación del historial? No se puede deshacer.",
                                  onConfirm: () => setLiquidaciones((list) => list.filter((x) => x.id !== r.id)),
                                })
                              }
                              style={{ padding: "7px 10px", borderRadius: 6, border: "1px solid #DADCD6", background: "#fff", color: "#A6432D", fontSize: 12.5, cursor: "pointer" }}
                            >
                              <Trash2 size={13} />
                            </button>
                          </div>
                        </div>
                      ))}
                  </div>
                )}
                {historialAbierto && (
                  <div>
                    <button onClick={() => setHistorialAbierto(null)} style={{ marginBottom: 14, padding: "7px 12px", borderRadius: 6, border: "1px solid #DADCD6", background: "#fff", fontSize: 12.5, cursor: "pointer" }}>
                      ← Volver al historial
                    </button>
                    <LiquidacionView
                      liq={historialAbierto.liq}
                      etiqueta={`${historialAbierto.tipo === "mensual" ? "" : "Complementaria "}${monthLabel(historialAbierto.year, historialAbierto.month)}`}
                      nombreCentralizada={catalogos.centralizadaNombre}
                      eeMap={historialAbierto.eeMap}
                      readOnly
                      onExport={() =>
                        exportLiquidacionExcel(
                          historialAbierto.liq,
                          `liquidacion_${monthLabel(historialAbierto.year, historialAbierto.month).replace(/\s|·/g, "_")}.xlsx`,
                          monthLabel(historialAbierto.year, historialAbierto.month),
                          historialAbierto.eeMap
                        )
                      }
                      onPrint={(filtro) => setLiqPrint({ liq: historialAbierto.liq, etiqueta: monthLabel(historialAbierto.year, historialAbierto.month), eeMap: historialAbierto.eeMap, filtro })}
                      confirmado
                      fechaConfirmacion={historialAbierto.fechaConfirmacion}
                    />
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* -------- CATÁLOGOS -------- */}
        {tab === "catalogos" && (
          <div>
            <div style={{ fontSize: 13, color: "#8A9088", marginBottom: 16, background: "#EAF3EF", padding: "10px 14px", borderRadius: 8 }}>
              Estos listados son compartidos: cualquier persona que acceda a esta herramienta ve y puede modificar los mismos valores.
            </div>

            {/* Habilitación centralizada */}
            <div style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, padding: 14, marginBottom: 16 }}>
              <div style={{ fontFamily: "'IBM Plex Serif', serif", fontWeight: 600, fontSize: 14, marginBottom: 8, color: "#1B2A4A" }}>Habilitación centralizada</div>
              <div style={{ fontSize: 12, color: "#8A9088", marginBottom: 8 }}>Nombre de la habilitación pagadora por la que se abona el gasto administrativo de las universidades marcadas como "centralizada" (ej. UNL). Asignación estímulo e IAPOS siempre los paga la habilitación pagadora del pasante.</div>
              <input
                style={{ width: "100%", maxWidth: 380, padding: "8px 10px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13.5 }}
                value={catalogos.centralizadaNombre}
                onChange={(e) => setCentralizadaNombre(e.target.value)}
                placeholder="Ej: Habilitación Pagadora Central"
              />
            </div>

            {/* Seguridad */}
            <div style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, padding: 14, marginBottom: 16 }}>
              <div style={{ fontFamily: "'IBM Plex Serif', serif", fontWeight: 600, fontSize: 14, marginBottom: 4, color: "#1B2A4A" }}>Seguridad</div>
              <div style={{ fontSize: 12, color: "#8A9088", marginBottom: 12 }}>
                Contraseñas opcionales para cargar/editar. Dejalas vacías si no querés pedir clave para esa acción. Cualquiera que entre a esta herramienta puede seguir <strong>consultando</strong> toda la información sin contraseña — esto solo protege agregar o modificar datos. Importante: es una protección simple pensada para coordinar el equipo, no un sistema de seguridad real (cualquiera con acceso al código puede eludirla).
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
                {[
                  ["pwdEdicion", "Cargar / editar pasantías y trámites"],
                  ["pwdLiquidacion", "Confirmar liquidaciones"],
                  ["pwdCatalogo", "Agregar en catálogos"],
                ].map(([key, label]) => (
                  <div key={key}>
                    <label style={{ fontSize: 11.5, fontWeight: 600, color: "#5B6158", marginBottom: 4, display: "block" }}>{label}</label>
                    <div style={{ position: "relative" }}>
                      <input
                        type={pwdVisible[key] ? "text" : "password"}
                        value={catalogos.seguridad?.[key] || ""}
                        onChange={(e) => setCatalogos((c) => ({ ...c, seguridad: { ...c.seguridad, [key]: e.target.value } }))}
                        placeholder="Sin contraseña"
                        autoComplete="new-password"
                        style={{ width: "100%", padding: "7px 32px 7px 9px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13, boxSizing: "border-box" }}
                      />
                      <button
                        type="button"
                        onClick={() => setPwdVisible((v) => ({ ...v, [key]: !v[key] }))}
                        title={pwdVisible[key] ? "Ocultar" : "Mostrar"}
                        style={{ position: "absolute", right: 6, top: "50%", transform: "translateY(-50%)", background: "none", border: "none", cursor: "pointer", color: "#8A9088", padding: 4, display: "flex" }}
                      >
                        {pwdVisible[key] ? <EyeOff size={15} /> : <Eye size={15} />}
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {/* Listas simples */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 16, marginBottom: 16 }}>
              {[
                ["jurisdicciones", "Jurisdicciones"],
                ["habilitaciones", "Habilitaciones pagadoras"],
                ["carreras", "Carreras"],
                ["estadosTramite", "Estados de trámite"],
              ].map(([key, label]) => (
                <div key={key} style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, padding: 14 }}>
                  <div style={{ fontFamily: "'IBM Plex Serif', serif", fontWeight: 600, fontSize: 14, marginBottom: 10, color: "#1B2A4A" }}>{label}</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 12 }}>
                    {catalogos[key].map((item) => (
                      <div key={item} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13, padding: "6px 9px", background: "#F6F7F5", borderRadius: 6 }}>
                        <span>{item}</span>
                        <button onClick={() => removeCatalogo(key, item)} style={{ background: "none", border: "none", color: "#A6432D", cursor: "pointer", padding: 2 }}>
                          <X size={13} />
                        </button>
                      </div>
                    ))}
                    {catalogos[key].length === 0 && <div style={{ fontSize: 12.5, color: "#8A9088" }}>Sin elementos todavía.</div>}
                  </div>
                  <div style={{ display: "flex", gap: 6 }}>
                    <input
                      placeholder="Agregar…"
                      value={newCatItem[key]}
                      onChange={(e) => setNewCatItem((n) => ({ ...n, [key]: e.target.value }))}
                      onKeyDown={(e) => {
                        if (e.key === "Enter" && newCatItem[key].trim()) {
                          ejecutarProtegido("catalogo", () => {
                            addCatalogo(key, newCatItem[key].trim());
                            setNewCatItem((n) => ({ ...n, [key]: "" }));
                          });
                        }
                      }}
                      style={{ flex: 1, padding: "7px 9px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13 }}
                    />
                    <button
                      onClick={() => {
                        if (newCatItem[key].trim()) {
                          ejecutarProtegido("catalogo", () => {
                            addCatalogo(key, newCatItem[key].trim());
                            setNewCatItem((n) => ({ ...n, [key]: "" }));
                          });
                        }
                      }}
                      style={{ padding: "7px 12px", borderRadius: 6, border: "none", background: "#1B2A4A", color: "#fff", fontSize: 13, cursor: "pointer" }}
                    >
                      <Plus size={14} />
                    </button>
                  </div>
                </div>
              ))}
            </div>

            {/* Universidades / Institutos */}
            <div style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, padding: 14, marginBottom: 16 }}>
              <div style={{ fontFamily: "'IBM Plex Serif', serif", fontWeight: 600, fontSize: 14, marginBottom: 10, color: "#1B2A4A" }}>Universidades / Institutos</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 12 }}>
                {catalogos.universidades.map((u) => (
                  <div key={u.nombre} style={{ display: "grid", gridTemplateColumns: "2fr 110px 140px auto", gap: 10, alignItems: "center", padding: "7px 9px", background: "#F6F7F5", borderRadius: 6, fontSize: 13 }}>
                    <span>{u.nombre}</span>
                    <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      <input
                        type="number"
                        value={u.gastoAdmin}
                        onChange={(e) => updateUniversidad(u.nombre, { gastoAdmin: Number(e.target.value) || 0 })}
                        style={{ width: 60, padding: "5px 7px", borderRadius: 5, border: "1px solid #DADCD6", fontSize: 12.5 }}
                      />
                      <span style={{ color: "#8A9088", fontSize: 12 }}>% admin.</span>
                    </div>
                    <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer" }}>
                      <input type="checkbox" checked={u.centralizada} onChange={(e) => updateUniversidad(u.nombre, { centralizada: e.target.checked })} />
                      Centralizada
                    </label>
                    <button onClick={() => removeUniversidad(u.nombre)} style={{ background: "none", border: "none", color: "#A6432D", cursor: "pointer", padding: 2, justifySelf: "end" }}>
                      <X size={13} />
                    </button>
                  </div>
                ))}
                {catalogos.universidades.length === 0 && <div style={{ fontSize: 12.5, color: "#8A9088" }}>Sin elementos todavía.</div>}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "2fr 110px 140px auto", gap: 10, alignItems: "center" }}>
                <input
                  placeholder="Nombre de la universidad / instituto"
                  value={nuevaUni.nombre}
                  onChange={(e) => setNuevaUni((n) => ({ ...n, nombre: e.target.value }))}
                  style={{ padding: "7px 9px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13 }}
                />
                <input
                  type="number"
                  placeholder="% admin."
                  value={nuevaUni.gastoAdmin}
                  onChange={(e) => setNuevaUni((n) => ({ ...n, gastoAdmin: e.target.value }))}
                  style={{ padding: "7px 9px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13 }}
                />
                <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, cursor: "pointer" }}>
                  <input type="checkbox" checked={nuevaUni.centralizada} onChange={(e) => setNuevaUni((n) => ({ ...n, centralizada: e.target.checked }))} />
                  Centralizada
                </label>
                <button onClick={() => ejecutarProtegido("catalogo", addUniversidad)} style={{ padding: "7px 12px", borderRadius: 6, border: "none", background: "#1B2A4A", color: "#fff", fontSize: 13, cursor: "pointer" }}>
                  <Plus size={14} />
                </button>
              </div>
            </div>

            {/* Asignación estímulo mensual */}
            <div style={{ background: "#fff", border: "1px solid #E7E9E4", borderRadius: 10, padding: 14 }}>
              <div style={{ fontFamily: "'IBM Plex Serif', serif", fontWeight: 600, fontSize: 14, marginBottom: 4, color: "#1B2A4A" }}>Asignación estímulo mensual</div>
              <div style={{ fontSize: 12, color: "#8A9088", marginBottom: 10 }}>
                Valor vigente por mes, según resolución del Ministerio de Economía. Es un historial: si un valor cambia en forma retroactiva (nueva norma), agregá el valor nuevo sin borrar el anterior — así las liquidaciones ya confirmadas con el valor viejo no se alteran, y podés armar la complementaria por la diferencia.
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 12 }}>
                {Object.entries(
                  catalogos.asignaciones.reduce((acc, a) => {
                    const key = `${a.year}-${a.month}`;
                    (acc[key] = acc[key] || []).push(a);
                    return acc;
                  }, {})
                )
                  .sort((a, b) => b[0].localeCompare(a[0]))
                  .map(([key, items]) => {
                    const ordenados = items.slice().sort((a, b) => (b.creadoEl || "").localeCompare(a.creadoEl || ""));
                    return (
                      <div key={key} style={{ border: "1px solid #E7E9E4", borderRadius: 8, overflow: "hidden" }}>
                        <div style={{ padding: "6px 9px", background: "#F6F7F5", fontSize: 12.5, fontWeight: 700, textTransform: "capitalize", color: "#1B2A4A" }}>
                          {monthLabel(ordenados[0].year, ordenados[0].month)}
                        </div>
                        {ordenados.map((a, i) => (
                          <div key={a.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 12.5, padding: "7px 9px", borderTop: i > 0 ? "1px solid #F0F1EE" : "none" }}>
                            <div>
                              <span style={{ fontWeight: 600 }}>{moneyFmt(a.monto)}</span>
                              {i === 0 && <span style={{ marginLeft: 6, fontSize: 10.5, padding: "1px 6px", borderRadius: 10, background: "#EAF3EF", color: "#2F6E5E" }}>vigente</span>}
                              {a.norma && <div style={{ color: "#8A9088", fontSize: 11 }}>Norma: {a.norma}</div>}
                              {a.observacion && <div style={{ color: "#8A9088", fontSize: 11 }}>{a.observacion}</div>}
                            </div>
                            <button onClick={() => removeAsignacionId(a.id)} style={{ background: "none", border: "none", color: "#A6432D", cursor: "pointer", padding: 2 }}>
                              <X size={13} />
                            </button>
                          </div>
                        ))}
                      </div>
                    );
                  })}
                {catalogos.asignaciones.length === 0 && <div style={{ fontSize: 12.5, color: "#8A9088" }}>Sin valores cargados todavía.</div>}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
                <select value={nuevaAsig.month} onChange={(e) => setNuevaAsig((n) => ({ ...n, month: Number(e.target.value) }))} style={{ padding: "7px 9px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13 }}>
                  {Array.from({ length: 12 }).map((_, i) => (
                    <option key={i} value={i}>{new Date(2000, i, 1).toLocaleDateString("es-AR", { month: "long" })}</option>
                  ))}
                </select>
                <select value={nuevaAsig.year} onChange={(e) => setNuevaAsig((n) => ({ ...n, year: Number(e.target.value) }))} style={{ padding: "7px 9px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13 }}>
                  {Array.from({ length: 7 }).map((_, i) => {
                    const y = new Date().getFullYear() - 2 + i;
                    return <option key={y} value={y}>{y}</option>;
                  })}
                </select>
                <input
                  type="number"
                  placeholder="Monto $"
                  value={nuevaAsig.monto}
                  onChange={(e) => setNuevaAsig((n) => ({ ...n, monto: e.target.value }))}
                  style={{ padding: "7px 9px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13 }}
                />
                <input
                  placeholder="N° de norma legal"
                  value={nuevaAsig.norma}
                  onChange={(e) => setNuevaAsig((n) => ({ ...n, norma: e.target.value }))}
                  style={{ padding: "7px 9px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13 }}
                />
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <input
                  placeholder="Observación (opcional, ej: reemplaza al valor anterior, ya se liquidaron las complementarias)"
                  value={nuevaAsig.observacion}
                  onChange={(e) => setNuevaAsig((n) => ({ ...n, observacion: e.target.value }))}
                  style={{ flex: 1, padding: "7px 9px", borderRadius: 6, border: "1px solid #DADCD6", fontSize: 13 }}
                />
                <button onClick={() => ejecutarProtegido("catalogo", addAsignacion)} style={{ padding: "7px 14px", borderRadius: 6, border: "none", background: "#1B2A4A", color: "#fff", fontSize: 13, cursor: "pointer", whiteSpace: "nowrap" }}>
                  <Plus size={14} /> Agregar
                </button>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Modal formulario */}
      {modal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,74,0.35)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "40px 20px", overflowY: "auto", zIndex: 50 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: 26, width: "100%", maxWidth: 640, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
              <div style={{ fontFamily: "'IBM Plex Serif', serif", fontWeight: 700, fontSize: 18, color: "#1B2A4A" }}>
                {modal.mode === "new" ? "Nueva pasantía" : "Editar pasantía"}
              </div>
              <button onClick={() => { setModal(null); setDniError(""); }} style={{ background: "none", border: "none", cursor: "pointer", color: "#8A9088" }}>
                <X size={20} />
              </button>
            </div>
            <PasantiaForm
              initial={modal.data}
              onSave={savePasantia}
              onCancel={() => { setModal(null); setDniError(""); }}
              catalogos={catalogos}
              onAddCatalogo={addCatalogo}
              dniError={dniError}
            />
          </div>
        </div>
      )}

      {/* Confirmación de borrado */}
      {confirmDelete && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,74,0.35)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 60 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: 22, width: 340, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 8 }}>¿Eliminar esta pasantía?</div>
            <div style={{ fontSize: 13, color: "#5B6158", marginBottom: 18 }}>Esta acción no se puede deshacer.</div>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
              <button onClick={() => setConfirmDelete(null)} style={{ padding: "8px 14px", borderRadius: 6, border: "1px solid #DADCD6", background: "#fff", fontSize: 13.5, cursor: "pointer" }}>Cancelar</button>
              <button onClick={() => deletePasantia(confirmDelete)} style={{ padding: "8px 14px", borderRadius: 6, border: "none", background: "#A6432D", color: "#fff", fontSize: 13.5, cursor: "pointer", fontWeight: 600 }}>Eliminar</button>
            </div>
          </div>
        </div>
      )}

      {/* Modal de trámite */}
      {tramiteModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,74,0.35)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "40px 20px", overflowY: "auto", zIndex: 50 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: 26, width: "100%", maxWidth: 640, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
              <div style={{ fontFamily: "'IBM Plex Serif', serif", fontWeight: 700, fontSize: 18, color: "#1B2A4A" }}>
                {tramiteModal.mode === "new" ? "Nuevo trámite" : "Editar trámite"}
              </div>
              <button onClick={() => setTramiteModal(null)} style={{ background: "none", border: "none", cursor: "pointer", color: "#8A9088" }}>
                <X size={20} />
              </button>
            </div>
            <TramiteForm initial={tramiteModal.data} onSave={saveTramite} onCancel={() => setTramiteModal(null)} catalogos={catalogos} onAddCatalogo={addCatalogo} />
          </div>
        </div>
      )}

      {confirmDeleteTramite && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,74,0.35)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 60 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: 22, width: 340, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 8 }}>¿Eliminar este trámite?</div>
            <div style={{ fontSize: 13, color: "#5B6158", marginBottom: 18 }}>Esta acción no se puede deshacer.</div>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
              <button onClick={() => setConfirmDeleteTramite(null)} style={{ padding: "8px 14px", borderRadius: 6, border: "1px solid #DADCD6", background: "#fff", fontSize: 13.5, cursor: "pointer" }}>Cancelar</button>
              <button onClick={() => deleteTramite(confirmDeleteTramite)} style={{ padding: "8px 14px", borderRadius: 6, border: "none", background: "#A6432D", color: "#fff", fontSize: 13.5, cursor: "pointer", fontWeight: 600 }}>Eliminar</button>
            </div>
          </div>
        </div>
      )}

      {presupuestoModal && (() => {
        const presupuesto = calcularPresupuestoTramite(presupuestoModal, catalogos);
        if (!presupuesto) return null;
        const th = { textAlign: "left", padding: "7px 8px", fontSize: 11.5, fontWeight: 700, color: "#5B6158", textTransform: "uppercase", borderBottom: "2px solid #E7E9E4" };
        const td = { padding: "7px 8px", fontSize: 12.5, borderBottom: "1px solid #F0F1EE" };
        return (
          <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,74,0.35)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "40px 20px", overflowY: "auto", zIndex: 50 }}>
            <div style={{ background: "#fff", borderRadius: 12, padding: 26, width: "100%", maxWidth: 640, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                <div style={{ fontFamily: "'IBM Plex Serif', serif", fontWeight: 700, fontSize: 17, color: "#1B2A4A" }}>Presupuesto estimado</div>
                <button onClick={() => setPresupuestoModal(null)} style={{ background: "none", border: "none", cursor: "pointer", color: "#8A9088" }}>
                  <X size={20} />
                </button>
              </div>
              <div style={{ fontSize: 12.5, color: "#5B6158", marginBottom: 16 }}>
                Expediente {presupuestoModal.expediente} · {presupuesto.cantidadPasantesTotal} pasante{presupuesto.cantidadPasantesTotal !== 1 ? "s" : ""} en total · desde {fmt(presupuestoModal.fechaProbableInicio)} hasta el 31/12/{presupuesto.detalleMeses[0]?.year}
              </div>
              {presupuesto.detalleMeses.map((d) => (
                <div key={`${d.year}-${d.month}`} style={{ marginBottom: 16 }}>
                  <div style={{ fontSize: 13, fontWeight: 700, color: "#1B2A4A", marginBottom: 6, textTransform: "capitalize" }}>
                    {monthLabel(d.year, d.month)}
                    {d.sinDato ? " — sin valor cargado" : d.estimado ? " (estimado)" : ""}
                    {d.esProrrateado && !d.sinDato && <span style={{ fontSize: 11, color: "#B8862F", fontWeight: 600 }}> · Proporcional ({d.diasTrabajados} días)</span>}
                  </div>
                  {d.sinDato ? (
                    <div style={{ fontSize: 12.5, color: "#A6432D" }}>No hay valor de asignación cargado para este mes.</div>
                  ) : (
                    <>
                      <table style={{ width: "100%", borderCollapse: "collapse" }}>
                        <thead>
                          <tr>
                            <th style={th}>Carrera</th>
                            <th style={th}>Universidad</th>
                            <th style={th}>Cant.</th>
                            <th style={th}>% Adm.</th>
                            <th style={th}>Total x pasante</th>
                            <th style={th}>Subtotal</th>
                          </tr>
                        </thead>
                        <tbody>
                          {d.porRenglon.map((r, i) => (
                            <tr key={i}>
                              <td style={td}>{r.carrera || "—"}</td>
                              <td style={td}>{r.universidad}</td>
                              <td style={td}>{r.cantidadPasantes}</td>
                              <td style={td}>{r.gastoPct}%</td>
                              <td style={td}>{moneyFmt(r.totalPorPasante)}</td>
                              <td style={{ ...td, fontWeight: 700 }}>{moneyFmt(r.subtotal)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                      <div style={{ textAlign: "right", fontSize: 12.5, marginTop: 4, color: "#1B2A4A" }}>Total del mes: <strong>{moneyFmt(d.totalMes)}</strong></div>
                    </>
                  )}
                </div>
              ))}
              <div style={{ fontSize: 11.5, color: "#8A9088", marginBottom: 14 }}>
                Los meses marcados "(estimado)" usan el último valor de asignación estímulo cargado en Catálogos, porque todavía no tenés el valor oficial de ese mes. Actualizá el cálculo cuando lo publiquen.
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={{ fontSize: 15, fontWeight: 700, color: "#1B2A4A" }}>Total: {moneyFmt(presupuesto.totalGeneral)}</div>
                <button
                  onClick={() => exportPresupuestoExcel(presupuestoModal, presupuesto)}
                  style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", background: "#1B2A4A", color: "#fff", border: "none", borderRadius: 7, fontSize: 13, fontWeight: 600, cursor: "pointer" }}
                >
                  <Download size={14} /> Excel
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {printData && <AvisoPDF items={printData} onClose={() => setPrintData(null)} />}
      {passwordPrompt && (
        <PasswordModal
          label={CATEGORIA_LABEL[passwordPrompt.categoria]}
          error={passwordError}
          onCancel={() => { setPasswordPrompt(null); setPasswordError(""); }}
          onSubmit={(value) => {
            const pwd = catalogos.seguridad?.[CATEGORIA_PWD_KEY[passwordPrompt.categoria]];
            if (value === pwd) {
              setDesbloqueado((d) => ({ ...d, [passwordPrompt.categoria]: true }));
              const accion = passwordPrompt.accion;
              setPasswordPrompt(null);
              setPasswordError("");
              accion();
            } else {
              setPasswordError("Contraseña incorrecta.");
            }
          }}
        />
      )}

      {confirmAction && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(27,42,74,0.35)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 80 }}>
          <div style={{ background: "#fff", borderRadius: 12, padding: 22, width: 360, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 13.5, color: "#20241F", marginBottom: 18, lineHeight: 1.5 }}>{confirmAction.message}</div>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
              <button onClick={() => setConfirmAction(null)} style={{ padding: "8px 14px", borderRadius: 6, border: "1px solid #DADCD6", background: "#fff", fontSize: 13.5, cursor: "pointer" }}>
                Cancelar
              </button>
              <button
                onClick={() => {
                  confirmAction.onConfirm();
                  setConfirmAction(null);
                }}
                style={{ padding: "8px 14px", borderRadius: 6, border: "none", background: "#1B2A4A", color: "#fff", fontSize: 13.5, cursor: "pointer", fontWeight: 600 }}
              >
                Confirmar
              </button>
            </div>
          </div>
        </div>
      )}

      {liqPrint && (
        <LiquidacionPrint
          liq={liqPrint.liq}
          etiqueta={liqPrint.etiqueta}
          nombreCentralizada={catalogos.centralizadaNombre}
          eeMap={liqPrint.eeMap}
          filtro={liqPrint.filtro}
          onClose={() => setLiqPrint(null)}
        />
      )}
    </div>
  );
}
