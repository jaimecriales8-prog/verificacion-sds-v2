'use client'

import { useState, useEffect, useRef, useCallback } from 'react'
import { supabase } from '@/lib/supabase'
import * as XLSX from 'xlsx'

const ADMIN_PASSWORD = process.env.NEXT_PUBLIC_ADMIN_PASSWORD || 'SDS2026'

type Ejecutor = { cedula: string; nombre: string; coordinador: string; lider: string; indice: number }
type Validacion = { cedula: string; estado: string }

type FilaLider = {
  coordinador: string
  lider: string
  total: number
  si: number
}

export default function AdminPage() {
  const [unlocked, setUnlocked] = useState(false)
  const [clave, setClave] = useState('')
  const [claveErr, setClaveErr] = useState('')
  const [validacionActiva, setValidacionActiva] = useState(true)
  const [ejecutores, setEjecutores] = useState<Ejecutor[]>([])
  const [validaciones, setValidaciones] = useState<Map<string, string>>(new Map())
  const [filtroCoord, setFiltroCoord] = useState('')
  const [filtroLider, setFiltroLider] = useState('')
  const [filtroEst, setFiltroEst] = useState('')
  const [alerta, setAlerta] = useState<{ tipo: 'ok' | 'er'; msg: string } | null>(null)
  const [cargando, setCargando] = useState(false)
  const fileRef = useRef<HTMLInputElement>(null)

  const cargarDatos = useCallback(async () => {
    const [{ data: ejs }, { data: vals }, { data: cfg }] = await Promise.all([
      supabase.from('ejecutores').select('*').order('indice'),
      supabase.from('validaciones').select('cedula,estado'),
      supabase.from('config').select('value').eq('key', 'validacion_activa').single(),
    ])
    if (ejs) setEjecutores(ejs)
    if (vals) setValidaciones(new Map(vals.map((v: Validacion) => [v.cedula, v.estado])))
    if (cfg) setValidacionActiva(cfg.value !== 'off')
  }, [])

  useEffect(() => {
    if (unlocked) {
      cargarDatos()
      const interval = setInterval(cargarDatos, 30000)
      return () => clearInterval(interval)
    }
  }, [unlocked, cargarDatos])

  function showAlerta(tipo: 'ok' | 'er', msg: string) {
    setAlerta({ tipo, msg })
    setTimeout(() => setAlerta(null), 5000)
  }

  function checkClave() {
    if (clave === ADMIN_PASSWORD) {
      setUnlocked(true)
      setClaveErr('')
    } else {
      setClaveErr('Clave incorrecta.')
      setClave('')
    }
  }

  async function toggleValidacion() {
    const nueva = !validacionActiva
    const msg = nueva ? '¿Encender la validación?' : '¿Apagar la validación? Los validadores no podrán ingresar documentos.'
    if (!confirm(msg)) return
    await supabase.from('config').upsert({ key: 'validacion_activa', value: nueva ? 'on' : 'off' })
    setValidacionActiva(nueva)
  }

  async function cargarExcel(file: File, modo: 'mantener' | 'reset') {
    setCargando(true)
    try {
      const buffer = await file.arrayBuffer()
      const wb = XLSX.read(buffer, { type: 'array' })
      const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes('consolid') || n.toLowerCase().includes('base') || n.toLowerCase().includes('datos')) || wb.SheetNames.find(n => {
        const ws2 = wb.Sheets[n]
        const rows2 = XLSX.utils.sheet_to_json<string[]>(ws2, { header: 1, defval: '' })
        const h = rows2[0] as string[]
        return h && h.includes('CEDULA') && h.includes('ROL')
      }) || wb.SheetNames[0]
      const ws = wb.Sheets[sheetName]
      const rows = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: '' })
      const hdr = rows[0] as string[]
      const iC = hdr.indexOf('CEDULA'), iN = hdr.indexOf('NOMBRES COMPLETOS')
      const iA = hdr.indexOf('APELLIDOS COMPLETO'), iR = hdr.indexOf('ROL')
      const iL = hdr.indexOf('LIDER')

      if (iC < 0 || iR < 0) {
        showAlerta('er', `Formato no reconocido. Columnas: ${hdr.slice(0, 8).join(', ')}`)
        return
      }

      const nuevos: Ejecutor[] = []
      let cc = '', cl = '', ri = 0

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i] as string[]
        const rol = String(row[iR] || '').trim()
        const ced = String(row[iC] || '').trim().replace(/\.0$/, '')
        const nom = iN >= 0 && iA >= 0
          ? `${row[iN] || ''} ${row[iA] || ''}`.trim()
          : String(row[iN >= 0 ? iN : iA] || '').trim()

        if (rol === 'COORDINADOR') { cc = nom; cl = '' }
        else if (rol === 'LIDER') {
          cl = nom || (iL >= 0 ? String(row[iL] || '').trim() : '')
        }
        else if (rol === 'EJECUTOR' && ced && ced !== 'nan') {
          nuevos.push({ cedula: ced, nombre: nom, coordinador: cc, lider: cl, indice: ri++ })
        }
      }

      if (nuevos.length === 0) { showAlerta('er', 'No se encontraron ejecutores en el archivo.'); return }

      // Borrar y reemplazar ejecutores
      await supabase.from('ejecutores').delete().neq('cedula', '__never__')
      const chunkSize = 500
      for (let i = 0; i < nuevos.length; i += chunkSize) {
        await supabase.from('ejecutores').insert(nuevos.slice(i, i + chunkSize))
      }

      if (modo === 'reset') {
        await supabase.from('validaciones').delete().neq('cedula', '__never__')
        setValidaciones(new Map())
      }

      setEjecutores(nuevos)
      showAlerta('ok', `✓ Base actualizada: ${nuevos.length.toLocaleString()} ejecutores.${modo === 'reset' ? ' Validaciones reiniciadas.' : ''}`)
    } catch (err: unknown) {
      showAlerta('er', `Error: ${err instanceof Error ? err.message : 'desconocido'}`)
    } finally {
      setCargando(false)
      if (fileRef.current) fileRef.current.value = ''
    }
  }

  function pedirModoYCargar(file: File) {
    const modo = confirm('¿Mantener validaciones existentes?\n\nAceptar = Mantener\nCancelar = Reiniciar todo')
      ? 'mantener' : 'reset'
    cargarExcel(file, modo)
  }

  async function descargar() {
    const rows = ejecutores.map(e => ({
      CEDULA: e.cedula,
      NOMBRE: e.nombre,
      COORDINADOR: e.coordinador,
      LIDER: e.lider,
      SI: validaciones.get(e.cedula) === 'SI' ? 'SI' : '',
    }))
    const ws = XLSX.utils.json_to_sheet(rows)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Validaciones')
    XLSX.writeFile(wb, `validaciones_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  const totalSI = validaciones.size
  const total = ejecutores.length

  // Construir tabla agrupada por lider
  const coordsSet = new Set(ejecutores.map(e => e.coordinador).filter(Boolean))
  const coords = Array.from(coordsSet)

  const ejFiltrados = ejecutores.filter(e => {
    if (filtroCoord && e.coordinador !== filtroCoord) return false
    if (filtroLider && e.lider !== filtroLider) return false
    if (filtroEst === 'SI' && validaciones.get(e.cedula) !== 'SI') return false
    if (filtroEst === 'PEND' && validaciones.get(e.cedula) === 'SI') return false
    return true
  })

  const lideresSet = new Set(
    ejecutores.filter(e => !filtroCoord || e.coordinador === filtroCoord).map(e => e.lider).filter(Boolean)
  )
  const lideres = Array.from(lideresSet)

  // Agrupar por lider para la tabla
  const porLider = new Map<string, FilaLider>()
  ejFiltrados.forEach(e => {
    const key = `${e.coordinador}||${e.lider}`
    if (!porLider.has(key)) porLider.set(key, { coordinador: e.coordinador, lider: e.lider, total: 0, si: 0 })
    const fila = porLider.get(key)!
    fila.total++
    if (validaciones.get(e.cedula) === 'SI') fila.si++
  })
  const filas = Array.from(porLider.values())

  if (!unlocked) {
    return (
      <div className="min-h-screen bg-[#fafaf8] flex items-center justify-center px-4">
        <div className="bg-white border border-[#e2e0d8] rounded-2xl p-8 w-full max-w-sm">
          <h1 className="text-lg font-bold text-[#085041] mb-1">Administración</h1>
          <p className="text-[13px] text-[#888780] mb-6">Ingresa la clave para continuar</p>
          <input
            type="password"
            placeholder="Clave..."
            value={clave}
            onChange={e => setClave(e.target.value)}
            onKeyDown={e => e.key === 'Enter' && checkClave()}
            className="w-full px-4 py-3 border-[1.5px] border-[#e2e0d8] rounded-xl mb-3 outline-none focus:border-[#1D9E75] text-[15px]"
            autoFocus
          />
          {claveErr && <p className="text-[#A32D2D] text-[12px] mb-3">{claveErr}</p>}
          <button onClick={checkClave} className="w-full py-3 bg-[#085041] text-white rounded-xl font-semibold">
            Entrar
          </button>
          <div className="mt-4 text-center">
            <a href="/" className="text-[12px] text-[#888780]">← Volver</a>
          </div>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-[#fafaf8]">
      {/* Topbar */}
      <div className="sticky top-0 z-10 bg-white border-b border-[#e2e0d8] px-4 py-3 flex items-center justify-between gap-3 flex-wrap">
        <span className="text-[15px] font-bold text-[#085041]">Administración</span>
        <div className="flex gap-2 flex-wrap items-center">
          <button
            onClick={toggleValidacion}
            className={`px-3 py-1.5 rounded-lg text-[12px] font-bold text-white ${validacionActiva ? 'bg-[#A32D2D]' : 'bg-[#1D9E75]'}`}
          >
            ⚡ {validacionActiva ? 'Apagar validación' : 'Encender validación'}
          </button>
          <label className={`px-3 py-1.5 rounded-lg text-[12px] font-bold text-white bg-[#534AB7] cursor-pointer ${cargando ? 'opacity-60' : ''}`}>
            {cargando ? 'Cargando...' : '📄 Cambiar base'}
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              disabled={cargando}
              onChange={e => { const f = e.target.files?.[0]; if (f) pedirModoYCargar(f) }}
            />
          </label>
          <button onClick={descargar} className="px-3 py-1.5 rounded-lg text-[12px] font-bold bg-[#e2e0d8] text-[#444441]">
            ↓ Excel
          </button>
          <button onClick={() => setUnlocked(false)} className="px-3 py-1.5 rounded-lg text-[12px] font-bold bg-[#444441] text-white">
            🔒 Bloquear
          </button>
        </div>
      </div>

      <div className="max-w-3xl mx-auto px-4 py-5">
        {alerta && (
          <div className={`rounded-xl px-4 py-3 mb-4 text-[13px] font-semibold ${alerta.tipo === 'ok' ? 'bg-[#E1F5EE] text-[#085041]' : 'bg-[#FCEBEB] text-[#A32D2D]'}`}>
            {alerta.msg}
          </div>
        )}

        {/* Stats */}
        <div className="grid grid-cols-3 gap-3 mb-5">
          <div className="bg-white border border-[#e2e0d8] rounded-xl p-4 text-center">
            <div className="text-[22px] font-bold text-[#085041]">{totalSI}</div>
            <div className="text-[11px] text-[#888780] mt-1">Confirmados SI</div>
            <div className="text-[11px] text-[#888780]">{total ? Math.round(totalSI / total * 100) : 0}%</div>
          </div>
          <div className="bg-white border border-[#e2e0d8] rounded-xl p-4 text-center">
            <div className="text-[22px] font-bold text-[#2C2C2A]">{total - totalSI}</div>
            <div className="text-[11px] text-[#888780] mt-1">Sin validar</div>
            <div className="text-[11px] text-[#888780]">{total ? Math.round((total - totalSI) / total * 100) : 0}%</div>
          </div>
          <div className="bg-white border border-[#e2e0d8] rounded-xl p-4 text-center">
            <div className={`text-[14px] font-bold mt-1 ${validacionActiva ? 'text-[#1D9E75]' : 'text-[#A32D2D]'}`}>
              {validacionActiva ? 'Activa' : 'Apagada'}
            </div>
            <div className="text-[11px] text-[#888780] mt-1">Validación</div>
            <div className="text-[11px] text-[#888780]">{total.toLocaleString()} ejecutores</div>
          </div>
        </div>

        {/* Filtros */}
        <div className="flex gap-2 flex-wrap mb-4">
          <select value={filtroCoord} onChange={e => { setFiltroCoord(e.target.value); setFiltroLider('') }}
            className="flex-1 min-w-[160px] px-3 py-2 border border-[#e2e0d8] rounded-lg text-[13px] bg-white">
            <option value="">Todos los coordinadores</option>
            {coords.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
          <select value={filtroLider} onChange={e => setFiltroLider(e.target.value)}
            className="flex-1 min-w-[160px] px-3 py-2 border border-[#e2e0d8] rounded-lg text-[13px] bg-white">
            <option value="">Todos los líderes</option>
            {lideres.map(l => <option key={l} value={l}>{l}</option>)}
          </select>
          <select value={filtroEst} onChange={e => setFiltroEst(e.target.value)}
            className="px-3 py-2 border border-[#e2e0d8] rounded-lg text-[13px] bg-white">
            <option value="">Todos</option>
            <option value="SI">Solo SI</option>
            <option value="PEND">Sin validar</option>
          </select>
        </div>

        {/* Tabla */}
        <div className="bg-white border border-[#e2e0d8] rounded-xl overflow-hidden">
          <table className="w-full text-[13px]">
            <thead>
              <tr className="border-b border-[#e2e0d8] text-[11px] text-[#888780] uppercase">
                <th className="text-left px-4 py-3 font-semibold">Líder</th>
                <th className="text-right px-4 py-3 font-semibold">Total</th>
                <th className="text-right px-4 py-3 font-semibold text-[#085041]">SI</th>
                <th className="text-right px-4 py-3 font-semibold">%</th>
              </tr>
            </thead>
            <tbody>
              {filas.length === 0 && (
                <tr><td colSpan={4} className="px-4 py-8 text-center text-[#888780]">Sin resultados</td></tr>
              )}
              {filas.map((f, i) => (
                <tr key={i} className="border-t border-[#f1efe8]">
                  <td className="px-4 py-3">
                    <div className="font-semibold text-[#2C2C2A]">{f.lider || '—'}</div>
                    <div className="text-[11px] text-[#888780]">{f.coordinador}</div>
                  </td>
                  <td className="px-4 py-3 text-right text-[#888780]">{f.total}</td>
                  <td className="px-4 py-3 text-right font-bold text-[#085041]">{f.si}</td>
                  <td className="px-4 py-3 text-right text-[#888780]">
                    {f.total ? Math.round(f.si / f.total * 100) : 0}%
                  </td>
                </tr>
              ))}
            </tbody>
            {filas.length > 0 && (
              <tfoot>
                <tr className="border-t-2 border-[#e2e0d8] bg-[#fafaf8] font-bold">
                  <td className="px-4 py-3 text-[12px]">Total</td>
                  <td className="px-4 py-3 text-right">{ejFiltrados.length}</td>
                  <td className="px-4 py-3 text-right text-[#085041]">
                    {ejFiltrados.filter(e => validaciones.get(e.cedula) === 'SI').length}
                  </td>
                  <td className="px-4 py-3 text-right">
                    {ejFiltrados.length ? Math.round(ejFiltrados.filter(e => validaciones.get(e.cedula) === 'SI').length / ejFiltrados.length * 100) : 0}%
                  </td>
                </tr>
              </tfoot>
            )}
          </table>
        </div>

        <div className="mt-6 text-center">
          <a href="/" className="text-[12px] text-[#888780]">← Vista validador</a>
        </div>
      </div>
    </div>
  )
}
