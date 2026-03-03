/** Toggle (on/off switch) */
export default function Toggle({ checked, onChange, label }) {
  return (
    <label className="flex items-center gap-2.5 cursor-pointer select-none">
      <span
        className={`toggle-track ${checked ? 'bg-brand-500' : 'bg-slate-300'}`}
        onClick={() => onChange(!checked)}
      >
        <span
          className={`toggle-thumb ${checked ? 'translate-x-4' : 'translate-x-0'}`}
        />
      </span>
      {label && <span className="text-sm text-slate-700">{label}</span>}
    </label>
  )
}
