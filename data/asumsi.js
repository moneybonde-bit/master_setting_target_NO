/* ==========================================================================
   tblAsumsi  —  data/asumsi.js   (5 indikator diatur khusus)
   --------------------------------------------------------------------------
   Arah dan Target Delta per indikator. Yang tidak terdaftar dianggap Naik dengan
   Target Delta = target_delta_default. Nama harus PERSIS sama dengan indicators.js.
   ========================================================================== */
window.WVI_ASUMSI = {
  target_delta_default: 0.1,
  rows: [
  {ind:"OIOS 160: Prevalence of underweight in children under five years of age", arah:"Turun", delta:-0.12},
  {ind:"Proportion of adolescent who married", arah:"Turun", delta:-0.1},
  {ind:"Proportion of adolescent who report L1 having experienced physical violence and /or psychological agression by parent / caregiver in the past 12 months", arah:"Turun", delta:-0.1},
  {ind:"L1 OIOS #45 Proportion of children in Grade 3 achieving at least a minimum proficiency level in reading.", arah:"Naik", delta:0.2},
  {ind:"OIOS 141: Proportion of households who have provided feedback to service providers, local governments, decision-makers, or other significant actors in relation to service delivery improvement", arah:"Naik", delta:0.2}
  ]
};
