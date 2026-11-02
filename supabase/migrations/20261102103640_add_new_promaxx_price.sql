-- Null means that this SKU has not had a saved R05.105 calculation yet.
alter table public.products
  add column if not exists new_promaxx_price numeric;

comment on column public.products.new_promaxx_price is
  'Latest Grab price calculated from R05.105 level 0; separate from Promaxx and the base selling price.';
