import { AlignLeft, AlignCenter, AlignRight } from 'lucide-react'
import Toggle from './Toggle.jsx'
import ColorPicker from './ColorPicker.jsx'
import SliderField from './SliderField.jsx'
import { FONT_OPTIONS, ALIGN_OPTIONS } from '../constants.js'

function Section({ title, children }) {
  return (
    <div className="space-y-4">
      <p className="text-[10px] font-bold uppercase tracking-widest text-slate-400 pb-0.5 border-b border-slate-100">
        {title}
      </p>
      {children}
    </div>
  )
}

export default function SongSettingsPanel({ settings, onChange }) {
  const set = (key, val) => onChange({ ...settings, [key]: val })

  return (
    <div className="space-y-6 text-sm">

      {/* ── Slides ── */}
      <Section title="Slide Layout">
        <div>
          <label className="label">Lines per Slide</label>
          <select
            className="select"
            value={settings.lines_per_slide}
            onChange={e => set('lines_per_slide', parseInt(e.target.value, 10))}
          >
            {[1, 2, 3, 4, 5, 6].map(n => (
              <option key={n} value={n}>{n} line{n > 1 ? 's' : ''}</option>
            ))}
          </select>
        </div>
      </Section>

      {/* ── Text ── */}
      <Section title="Text">
        <div>
          <label className="label">Font Family</label>
          <select
            className="select"
            value={settings.font_family}
            onChange={e => set('font_family', e.target.value)}
          >
            {FONT_OPTIONS.map(f => (
              <option key={f} value={f}>{f}</option>
            ))}
          </select>
        </div>

        <SliderField
          label="Font Size"
          value={settings.font_size}
          onChange={v => set('font_size', v)}
          min={12} max={100} step={1} unit=" pt"
        />

        <Toggle
          label="Bold"
          checked={settings.bold}
          onChange={v => set('bold', v)}
        />

        <div>
          <label className="label">Alignment</label>
          <div className="flex gap-1.5">
            {ALIGN_OPTIONS.map(({ value, label }) => (
              <button
                key={value}
                onClick={() => set('align', value)}
                title={label}
                className={`flex-1 py-2 rounded-lg border text-xs font-medium flex items-center justify-center gap-1 transition
                  ${settings.align === value
                    ? 'bg-brand-500 text-white border-brand-500'
                    : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'}`}
              >
                {value === 'left'   && <AlignLeft  size={13} />}
                {value === 'center' && <AlignCenter size={13} />}
                {value === 'right'  && <AlignRight  size={13} />}
                <span>{label}</span>
              </button>
            ))}
          </div>
        </div>

        <SliderField
          label="Line Spacing"
          value={settings.line_spacing}
          onChange={v => set('line_spacing', v)}
          min={1.0} max={3.5} step={0.1}
        />

        <SliderField
          label="Verse Spacing"
          value={settings.verse_spacing}
          onChange={v => set('verse_spacing', v)}
          min={0.0} max={3.0} step={0.1}
        />
      </Section>

      {/* ── Colors ── */}
      <Section title="Colors">
        <ColorPicker
          label="Background"
          value={settings.bg_color}
          onChange={v => set('bg_color', v)}
        />
        <ColorPicker
          label="Lyrics Color"
          value={settings.text_color}
          onChange={v => set('text_color', v)}
        />
        <ColorPicker
          label="Song Title Color"
          value={settings.ref_color}
          onChange={v => set('ref_color', v)}
        />
      </Section>

      {/* ── Song Title ── */}
      <Section title="Song Title">
        <SliderField
          label="Title Font Size"
          value={settings.ref_font_size}
          onChange={v => set('ref_font_size', v)}
          min={8} max={48} step={1} unit=" pt"
        />
      </Section>

      {/* ── Layout ── */}
      <Section title="Layout">
        <SliderField
          label="Slide Padding"
          value={settings.padding}
          onChange={v => set('padding', v)}
          min={0.1} max={1.5} step={0.05}
          unit=" in"
        />
      </Section>

    </div>
  )
}
