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
  const [modalBase, setModalBase] = useState(false)
  const [archivoSeleccionado, setArchivoSeleccionado] = useState<File | null>(null)
  const [modalAjuste, setModalAjuste] = useState<Ejecutor | null>(null)
  const [ajusteForm, setAjusteForm] = useState({ cedula: '', nombre: '', coordinador: '', lider: '' })
  const [busquedaAjuste, setBusquedaAjuste] = useState('')
  const [tabActiva, setTabActiva] = useState<'consolidado' | 'ajustes'>('consolidado')
  const [coordAbiertos, setCoordAbiertos] = useState<Set<string>>(new Set())
  const fileRef = useRef<HTMLInputElement>(null)

  const cargarDatos = useCallback(async () => {
    // Paginar ejecutores de 1000 en 1000
    let todosEjs: Ejecutor[] = []
    let desde = 0
    const PAGE = 1000
    while (true) {
      const { data } = await supabase.from('ejecutores').select('*').order('indice').range(desde, desde + PAGE - 1)
      if (!data || data.length === 0) break
      todosEjs = todosEjs.concat(data)
      if (data.length < PAGE) break
      desde += PAGE
    }

    // Paginar validaciones
    let todasVals: Validacion[] = []
    desde = 0
    while (true) {
      const { data } = await supabase.from('validaciones').select('cedula,estado').range(desde, desde + PAGE - 1)
      if (!data || data.length === 0) break
      todasVals = todasVals.concat(data)
      if (data.length < PAGE) break
      desde += PAGE
    }

    const { data: cfg } = await supabase.from('config').select('value').eq('key', 'validacion_activa').single()
    setEjecutores(todosEjs)
    setValidaciones(new Map(todasVals.map((v: Validacion) => [v.cedula, v.estado])))
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

      // Mapa para deduplicar: última aparición de cada cédula gana
      const mapaEjs = new Map<string, Omit<Ejecutor, 'indice'>>()
      let cc = '', cl = ''

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i] as string[]
        const rol = String(row[iR] || '').trim()
        const ced = String(row[iC] || '').trim().replace(/\.0$/, '')
        const nom = iN >= 0 && iA >= 0
          ? `${row[iN] || ''} ${row[iA] || ''}`.trim()
          : String(row[iN >= 0 ? iN : iA] || '').trim()

        if (!ced || ced === 'nan' || !nom) continue

        // Actualizar contexto jerárquico
        if (rol === 'COORDINADOR') { cc = nom; cl = '' }
        else if (rol === 'LIDER') { cl = nom }

        // Guardar persona con su contexto actual (sobreescribe si ya existe)
        mapaEjs.set(ced, { cedula: ced, nombre: nom, coordinador: cc, lider: cl })
      }

      const nuevos: Ejecutor[] = Array.from(mapaEjs.values()).map((e, idx) => ({ ...e, indice: idx }))

      if (nuevos.length === 0) { showAlerta('er', 'No se encontraron ejecutores en el archivo.'); return }

      // Borrar en bucle con chunks de 200 cédulas
      let siguen = true
      while (siguen) {
        const { data: batch } = await supabase.from('ejecutores').select('cedula').limit(200)
        if (!batch || batch.length === 0) break
        await supabase.from('ejecutores').delete().in('cedula', batch.map(r => r.cedula))
        siguen = batch.length === 200
      }

      // Deduplicar por cédula (quedar con el primero)
      const vistos = new Set<string>()
      const unicos = nuevos.filter(e => { if (vistos.has(e.cedula)) return false; vistos.add(e.cedula); return true })

      // Upsert en chunks de 200
      for (let i = 0; i < unicos.length; i += 200) {
        const { error } = await supabase.from('ejecutores').upsert(unicos.slice(i, i + 200))
        if (error) { showAlerta('er', `Error en inserción (fila ${i}): ${error.message}`); return }
      }
      // Actualizar nuevos para el mensaje final
      nuevos.length = 0; unicos.forEach(u => nuevos.push(u))

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
    setArchivoSeleccionado(file)
    setModalBase(true)
  }

  function abrirAjuste(ej: Ejecutor) {
    setAjusteForm({ cedula: ej.cedula, nombre: ej.nombre, coordinador: ej.coordinador, lider: ej.lider })
    setModalAjuste(ej)
  }

  async function guardarAjuste() {
    if (!modalAjuste) return
    const { cedula, nombre, coordinador, lider } = ajusteForm
    const cedulaNueva = cedula.trim()
    const updates: Partial<Ejecutor> = { nombre: nombre.trim(), coordinador: coordinador.trim(), lider: lider.trim() }

    // Si cambió la cédula, insertar nuevo y borrar viejo
    if (cedulaNueva !== modalAjuste.cedula) {
      await supabase.from('ejecutores').insert({ ...modalAjuste, ...updates, cedula: cedulaNueva })
      await supabase.from('ejecutores').delete().eq('cedula', modalAjuste.cedula)
      // Migrar validación si existe
      const val = validaciones.get(modalAjuste.cedula)
      if (val) {
        await supabase.from('validaciones').upsert({ cedula: cedulaNueva, estado: val, updated_at: new Date().toISOString() })
        await supabase.from('validaciones').delete().eq('cedula', modalAjuste.cedula)
      }
    } else {
      await supabase.from('ejecutores').update(updates).eq('cedula', modalAjuste.cedula)
    }

    setModalAjuste(null)
    showAlerta('ok', '✓ Ejecutor actualizado')
    cargarDatos()
  }

  async function ajustarValidacion(cedula: string, accion: 'SI' | 'quitar') {
    if (accion === 'SI') {
      await supabase.from('validaciones').upsert({ cedula, estado: 'SI', updated_at: new Date().toISOString() })
      setValidaciones(prev => new Map(prev).set(cedula, 'SI'))
    } else {
      await supabase.from('validaciones').delete().eq('cedula', cedula)
      setValidaciones(prev => { const m = new Map(prev); m.delete(cedula); return m })
    }
  }

  function confirmarCambioBase(modo: 'mantener' | 'reset') {
    setModalBase(false)
    if (archivoSeleccionado) cargarExcel(archivoSeleccionado, modo)
    setArchivoSeleccionado(null)
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

        {/* Tabs */}
        <div className="flex gap-1 mb-5 bg-[#F1EFE8] rounded-xl p-1">
          <button onClick={() => setTabActiva('consolidado')}
            className={`flex-1 py-2 rounded-lg text-[13px] font-semibold transition-colors ${tabActiva === 'consolidado' ? 'bg-white text-[#085041] shadow-sm' : 'text-[#888780]'}`}>
            Consolidado
          </button>
          <button onClick={() => setTabActiva('ajustes')}
            className={`flex-1 py-2 rounded-lg text-[13px] font-semibold transition-colors ${tabActiva === 'ajustes' ? 'bg-white text-[#085041] shadow-sm' : 'text-[#888780]'}`}>
            Ajustes
          </button>
        </div>

        {/* TAB AJUSTES */}
        {tabActiva === 'ajustes' && (
          <div>
            <div className="relative mb-4">
              <input
                type="text"
                placeholder="Buscar por cédula o nombre..."
                value={busquedaAjuste}
                onChange={e => setBusquedaAjuste(e.target.value)}
                className="w-full px-4 py-3 pr-10 border-[1.5px] border-[#e2e0d8] rounded-xl outline-none focus:border-[#1D9E75] text-[14px] bg-white"
              />
              {busquedaAjuste && (
                <button onClick={() => setBusquedaAjuste('')}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-[#888780] text-xl leading-none">×</button>
              )}
            </div>
            <div className="bg-white border border-[#e2e0d8] rounded-xl overflow-hidden">
              {ejecutores
                .filter(e => {
                  const q = busquedaAjuste.toLowerCase()
                  return !q || e.cedula.includes(q) || e.nombre.toLowerCase().includes(q)
                })
                .slice(0, 50)
                .map(e => (
                  <div key={e.cedula} className="flex items-center justify-between px-4 py-3 border-b border-[#f1efe8] last:border-0">
                    <div>
                      <div className="text-[13px] font-semibold">{e.nombre}</div>
                      <div className="text-[11px] text-[#888780]">CC {e.cedula} · {e.lider}</div>
                    </div>
                    <div className="flex items-center gap-2">
                      <span className={`text-[11px] font-bold px-2 py-1 rounded-lg ${validaciones.get(e.cedula) === 'SI' ? 'bg-[#E1F5EE] text-[#085041]' : 'bg-[#F1EFE8] text-[#888780]'}`}>
                        {validaciones.get(e.cedula) === 'SI' ? 'SI' : 'PEND'}
                      </span>
                      <button onClick={() => abrirAjuste(e)}
                        className="text-[12px] px-3 py-1.5 bg-[#E6F1FB] text-[#185FA5] rounded-lg font-semibold">
                        Editar
                      </button>
                    </div>
                  </div>
                ))}
              {busquedaAjuste.length < 2 && (
                <div className="px-4 py-6 text-center text-[12px] text-[#888780]">Escribe al menos 2 caracteres para buscar</div>
              )}
            </div>
          </div>
        )}

        {/* TAB CONSOLIDADO */}
        {tabActiva === 'consolidado' && <>
        {/* Filtros */}
        <div className="flex gap-2 flex-wrap mb-4 items-center">
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
          {(filtroCoord || filtroLider || filtroEst) && (
            <button onClick={() => { setFiltroCoord(''); setFiltroLider(''); setFiltroEst('') }}
              className="px-3 py-2 rounded-lg text-[12px] font-semibold bg-[#F1EFE8] text-[#444441] whitespace-nowrap">
              × Limpiar
            </button>
          )}
        </div>

        {/* Tabla jerárquica acordeón */}
        {(() => {
          const coordsAgrupados = new Map<string, FilaLider[]>()
          filas.forEach(f => {
            if (!coordsAgrupados.has(f.coordinador)) coordsAgrupados.set(f.coordinador, [])
            coordsAgrupados.get(f.coordinador)!.push(f)
          })
          const totFil = ejFiltrados.length
          const siFil = ejFiltrados.filter(e => validaciones.get(e.cedula) === 'SI').length
          const pctFil = totFil ? Math.round(siFil / totFil * 100) : 0

          return (
            <div className="flex flex-col gap-2">
              {filas.length === 0 && (
                <div className="bg-white border border-[#e2e0d8] rounded-xl px-4 py-8 text-center text-[#888780] text-[13px]">Sin resultados</div>
              )}
              {Array.from(coordsAgrupados.entries()).map(([coord, liders]) => {
                const totCoord = liders.reduce((a, f) => a + f.total, 0)
                const siCoord = liders.reduce((a, f) => a + f.si, 0)
                const pctCoord = totCoord ? Math.round(siCoord / totCoord * 100) : 0
                const abierto = coordAbiertos.has(coord)
                const toggle = () => setCoordAbiertos(prev => {
                  const s = new Set(prev)
                  s.has(coord) ? s.delete(coord) : s.add(coord)
                  return s
                })
                return (
                  <div key={coord} className="bg-white border border-[#e2e0d8] rounded-xl overflow-hidden">
                    {/* Header coordinador */}
                    <button onClick={toggle} className="w-full flex items-center justify-between px-4 py-3 bg-[#E1F5EE] hover:bg-[#c5e8da] transition-colors">
                      <div className="flex items-center gap-2">
                        <span className="text-[13px] font-bold text-[#085041]">{coord || '—'}</span>
                        <span className={`text-[11px] transition-transform ${abierto ? 'rotate-180' : ''}`}>▼</span>
                      </div>
                      <div className="flex gap-4 text-[12px] font-semibold">
                        <span className="text-[#085041]">✓ {siCoord}</span>
                        <span className="text-[#A32D2D]">⏳ {totCoord - siCoord}</span>
                        <span className="text-[#085041]">{pctCoord}%</span>
                        <span className="text-[#888780] font-normal">{totCoord} ejec.</span>
                      </div>
                    </button>

                    {/* Líderes — solo si está abierto */}
                    {abierto && (
                      <table className="w-full text-[13px]">
                        <thead>
                          <tr className="border-b border-[#f1efe8] text-[11px] text-[#888780] uppercase bg-[#fafaf8]">
                            <th className="text-left px-4 py-2 font-semibold pl-8">Líder</th>
                            <th className="text-right px-4 py-2 font-semibold">Ejec.</th>
                            <th className="text-right px-4 py-2 font-semibold text-[#085041]">SI</th>
                            <th className="text-right px-4 py-2 font-semibold text-[#A32D2D]">Pend.</th>
                            <th className="text-right px-4 py-2 font-semibold">%</th>
                          </tr>
                        </thead>
                        <tbody>
                          {liders.map((f, i) => {
                            const pend = f.total - f.si
                            const pct = f.total ? Math.round(f.si / f.total * 100) : 0
                            return (
                              <tr key={i} className="border-t border-[#f1efe8] hover:bg-[#fafaf8]">
                                <td className="px-4 py-2.5 pl-8 text-[#2C2C2A]">{f.lider || '—'}</td>
                                <td className="px-4 py-2.5 text-right text-[#888780]">{f.total}</td>
                                <td className="px-4 py-2.5 text-right font-semibold text-[#085041]">{f.si}</td>
                                <td className="px-4 py-2.5 text-right text-[#A32D2D]">{pend}</td>
                                <td className="px-4 py-2.5 text-right text-[#888780]">{pct}%</td>
                              </tr>
                            )
                          })}
                        </tbody>
                      </table>
                    )}
                  </div>
                )
              })}

              {/* Total general */}
              {filas.length > 0 && (
                <div className="bg-[#fafaf8] border border-[#e2e0d8] rounded-xl px-4 py-3 flex justify-between text-[13px] font-bold">
                  <span>TOTAL</span>
                  <div className="flex gap-4">
                    <span className="text-[#085041]">✓ {siFil}</span>
                    <span className="text-[#A32D2D]">⏳ {totFil - siFil}</span>
                    <span>{pctFil}%</span>
                    <span className="text-[#888780] font-normal">{totFil} ejec.</span>
                  </div>
                </div>
              )}
            </div>
          )
        })()}

        </> }

        <div className="mt-6 text-center">
          <a href="/" className="text-[12px] text-[#888780]">← Vista validador</a>
        </div>
      </div>

      {/* Modal editar ejecutor */}
      {modalAjuste && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center px-4">
          <div className="bg-white rounded-2xl p-6 w-full max-w-sm shadow-xl">
            <div className="flex items-center justify-between mb-5">
              <h2 className="text-[16px] font-bold">✏️ Editar ejecutor</h2>
              <button onClick={() => setModalAjuste(null)} className="text-[#888780] text-xl px-1">×</button>
            </div>
            <div className="flex flex-col gap-3 mb-4">
              {[
                { label: 'Cédula', key: 'cedula' },
                { label: 'Nombre completo', key: 'nombre' },
                { label: 'Coordinador', key: 'coordinador' },
                { label: 'Líder', key: 'lider' },
              ].map(({ label, key }) => (
                <div key={key}>
                  <label className="text-[11px] text-[#888780] font-semibold uppercase mb-1 block">{label}</label>
                  <input
                    value={ajusteForm[key as keyof typeof ajusteForm]}
                    onChange={e => setAjusteForm(f => ({ ...f, [key]: e.target.value }))}
                    className="w-full px-3 py-2.5 border-[1.5px] border-[#e2e0d8] rounded-xl text-[14px] outline-none focus:border-[#1D9E75]"
                  />
                </div>
              ))}
            </div>

            {/* Ajuste de validación */}
            <div className="bg-[#F1EFE8] rounded-xl p-3 mb-4">
              <div className="text-[11px] font-bold text-[#888780] uppercase mb-2">Validación</div>
              <div className="flex gap-2">
                <button onClick={() => { ajustarValidacion(modalAjuste.cedula, 'SI'); setModalAjuste(null) }}
                  className="flex-1 py-2 bg-[#1D9E75] text-white rounded-lg text-[13px] font-bold">
                  Confirmar SI
                </button>
                <button onClick={() => { ajustarValidacion(modalAjuste.cedula, 'quitar'); setModalAjuste(null) }}
                  className="flex-1 py-2 border border-[#A32D2D] text-[#A32D2D] rounded-lg text-[13px] font-bold">
                  Quitar SI
                </button>
              </div>
            </div>

            <div className="flex gap-2">
              <button onClick={() => setModalAjuste(null)}
                className="flex-1 py-2.5 border border-[#e2e0d8] rounded-xl text-[13px] font-semibold text-[#5F5E5A]">
                Cancelar
              </button>
              <button onClick={guardarAjuste}
                className="flex-1 py-2.5 bg-[#085041] text-white rounded-xl text-[13px] font-bold">
                Guardar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal cambiar base */}
      {modalBase && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center px-4">
          <div className="bg-white rounded-2xl p-6 w-full max-w-sm shadow-xl">
            <div className="flex items-center justify-between mb-1">
              <h2 className="text-[16px] font-bold">📄 Cambiar base de datos</h2>
              <button onClick={() => { setModalBase(false); setArchivoSeleccionado(null); if (fileRef.current) fileRef.current.value = '' }}
                className="text-[#888780] text-xl leading-none px-1">×</button>
            </div>
            <p className="text-[13px] text-[#5F5E5A] mb-5">Selecciona cómo cargar la nueva base:</p>
            <div className="flex flex-col gap-3 mb-5">
              <button onClick={() => confirmarCambioBase('mantener')}
                className="border-[1.5px] border-[#e2e0d8] rounded-xl p-4 text-left hover:border-[#1D9E75]">
                <div className="text-[14px] font-bold mb-1">✓ Mantener validaciones</div>
                <div className="text-[12px] text-[#5F5E5A]">Carga la nueva base y conserva los SI ya registrados para las cédulas que coincidan.</div>
              </button>
              <button onClick={() => confirmarCambioBase('reset')}
                className="border-[1.5px] border-[#e2e0d8] rounded-xl p-4 text-left hover:border-[#A32D2D]">
                <div className="text-[14px] font-bold mb-1">⚠ Reiniciar todo</div>
                <div className="text-[12px] text-[#5F5E5A]">Carga la nueva base y borra todos los SI. Empieza desde cero.</div>
              </button>
            </div>
            <button onClick={() => { setModalBase(false); setArchivoSeleccionado(null); if (fileRef.current) fileRef.current.value = '' }}
              className="w-full py-2.5 border border-[#e2e0d8] rounded-xl text-[13px] font-semibold text-[#5F5E5A]">
              Cancelar
            </button>
          </div>
        </div>
      )}
    </div>
  )
}
