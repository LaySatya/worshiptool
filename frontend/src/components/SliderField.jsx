/** SliderField — labelled range slider with a numeric readout */
export default function SliderField({ label, value, onChange, min, max, step = 1, unit = '' }) {
  return (
    <div>
      <div className="flex justify-between items-center mb-1.5">
        <label className="label mb-0">{label}</label>
        <span className="text-xs font-semibold text-brand-500 tabular-nums">
          {typeof value === 'number' ? value.toFixed(step < 1 ? 1 : 0) : value}{unit}
        </span>
      </div>
      <input
        type="range"
        min={min}
        max={max}
        step={step}
        value={value}
        onChange={e => onChange(step < 1 ? parseFloat(e.target.value) : parseInt(e.target.value, 10))}
        className="slider"
      />
    </div>
  )
}
