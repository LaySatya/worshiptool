/** ColorPicker — native color input with a hex label */
export default function ColorPicker({ label, value, onChange }) {
  return (
    <div>
      <label className="label">{label}</label>
      <div className="flex items-center gap-2.5">
        <div className="relative">
          <input
            type="color"
            value={value}
            onChange={e => onChange(e.target.value)}
            className="w-9 h-9 rounded-lg border border-slate-200 cursor-pointer p-0.5 bg-white"
          />
        </div>
        <input
          type="text"
          value={value}
          onChange={e => {
            const v = e.target.value
            if (/^#[0-9a-fA-F]{0,6}$/.test(v)) onChange(v)
          }}
          maxLength={7}
          className="input w-28 font-mono text-xs"
        />
      </div>
    </div>
  )
}
