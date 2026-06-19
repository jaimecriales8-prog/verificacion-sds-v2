'use client'

import { useState, useEffect, useRef } from 'react'
import { supabase } from '@/lib/supabase'

function maskName(name: string) {
  return name.split(' ').map(w => {
    if (w.length <= 2) return w
    return w.slice(0, 2) + w.slice(2).replace(/./g, 'X')
  }).join(' ')
}

type Ejecutor = { nombre: string; coordinador: string; lider: string }

export default function ValidadorPage() {
  const [validacionActiva, setValidacionActiva] = useState(true)
  const [cedula, setCedula] = useState('')
  const [resultado, setResultado] = useState<{ ejecutor: Ejecutor; validado: boolean } | null>(null)
  const [alerta, setAlerta] = useState<{ tipo: 'ok' | 'er'; msg: string } | null>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    supabase.from('config').select('value').eq('key', 'validacion_activa').single()
      .then(({ data }) => { if (data) setValidacionActiva(data.value !== 'off') })
  }, [])

  function showAlerta(tipo: 'ok' | 'er', msg: string) {
    setAlerta({ tipo, msg })
    setTimeout(() => setAlerta(null), 4000)
  }

  async function buscar() {
    const ced = cedula.trim()
    setResultado(null)
    setAlerta(null)
    if (!validacionActiva) { showAlerta('er', 'La validación está desactivada.'); return }
    if (!ced) { showAlerta('er', 'Ingresa un número de cédula.'); return }

    const { data: ej } = await supabase.from('ejecutores').select('nombre,coordinador,lider').eq('cedula', ced).single()
    if (!ej) { showAlerta('er', `Cédula ${ced} no encontrada en la base.`); return }

    const { data: val } = await supabase.from('validaciones').select('estado').eq('cedula', ced).single()
    setResultado({ ejecutor: ej, validado: !!val })
  }

  async function confirmar() {
    if (!resultado) return
    const ced = cedula.trim()
    await supabase.from('validaciones').upsert({ cedula: ced, estado: 'SI', updated_at: new Date().toISOString() })
    showAlerta('ok', `✓ ${maskName(resultado.ejecutor.nombre)} confirmado como SI`)
    setResultado(null)
    setCedula('')
    inputRef.current?.focus()
  }

  async function anular() {
    if (!resultado || !confirm('¿Anular la validación de este ejecutor?')) return
    const ced = cedula.trim()
    await supabase.from('validaciones').delete().eq('cedula', ced)
    showAlerta('ok', `Validación de ${maskName(resultado.ejecutor.nombre)} anulada`)
    setResultado(null)
    setCedula('')
    inputRef.current?.focus()
  }

  return (
    <div className="min-h-screen bg-[#fafaf8]">
      <div className="max-w-md mx-auto px-5 pt-8 pb-16">
        <h1 className="text-lg font-bold text-[#085041] mb-6">Verificación SDS 2026</h1>

        {!validacionActiva && (
          <div className="bg-[#444441] text-white rounded-xl p-5 text-center mb-6">
            <p className="font-semibold text-[15px]">🔒 Validación cerrada</p>
            <span className="text-[13px] opacity-80">El ingreso de documentos ha sido desactivado</span>
          </div>
        )}

        {validacionActiva && (
          <>
            {alerta && (
              <div className={`rounded-xl px-4 py-3 mb-4 text-[13px] font-semibold ${alerta.tipo === 'ok' ? 'bg-[#E1F5EE] text-[#085041]' : 'bg-[#FCEBEB] text-[#A32D2D]'}`}>
                {alerta.msg}
              </div>
            )}

            <div className="flex gap-2 mb-4">
              <input
                ref={inputRef}
                type="text"
                inputMode="numeric"
                maxLength={12}
                placeholder="Número de cédula..."
                value={cedula}
                onChange={e => setCedula(e.target.value.replace(/\D/g, ''))}
                onKeyDown={e => e.key === 'Enter' && buscar()}
                className="flex-1 px-4 py-3 border-[1.5px] border-[#e2e0d8] rounded-xl text-lg font-medium bg-white outline-none focus:border-[#1D9E75]"
              />
              <button onClick={buscar} className="px-5 py-3 bg-[#1D9E75] text-white font-semibold rounded-xl active:opacity-80">
                Buscar
              </button>
            </div>

            {resultado && (
              <div className="bg-white border-[1.5px] border-[#e2e0d8] rounded-2xl p-5">
                <div className="text-[17px] font-bold mb-1">{maskName(resultado.ejecutor.nombre)}</div>
                <div className="text-[12px] text-[#888780] mb-3">CC {cedula}</div>

                {resultado.validado ? (
                  <>
                    <div className="flex items-center gap-2 px-3 py-2 rounded-lg bg-[#E1F5EE] text-[#085041] text-[13px] font-semibold mb-3">
                      ✓ Ya confirmado como SI
                    </div>
                    <button onClick={anular} className="w-full py-3 border-[1.5px] border-[#A32D2D] text-[#A32D2D] rounded-xl text-[14px] font-bold">
                      Anular validación
                    </button>
                  </>
                ) : (
                  <>
                    <div className="flex items-center gap-2 px-3 py-2 rounded-lg bg-[#F1EFE8] text-[#444441] text-[13px] font-semibold mb-3">
                      Sin validar
                    </div>
                    <button onClick={confirmar} className="w-full py-4 bg-[#085041] text-white rounded-xl text-[15px] font-bold active:opacity-80">
                      Confirmar SI
                    </button>
                  </>
                )}
              </div>
            )}
          </>
        )}

        <div className="mt-8 text-center">
          <a href="/admin" className="text-[12px] text-[#888780]">Administración</a>
        </div>
      </div>
    </div>
  )
}
