create table if not exists public.hourly_metrics (
  row_key text primary key,
  logged_at_utc timestamptz not null,
  logged_at_local text,
  sales_amount numeric,
  orders numeric,
  ad_spend numeric,
  roas numeric,
  source_sales text,
  source_marketing text,
  ingested_at_utc timestamptz not null default now()
);

create index if not exists hourly_metrics_logged_at_utc_idx on public.hourly_metrics (logged_at_utc);

create table if not exists public.shopify_orders (
  order_id text primary key,
  order_name text,
  created_at_utc timestamptz not null,
  processed_at_utc timestamptz,
  currency text,
  source_name text,
  customer_id text,
  customer_type text,
  gross_sales numeric,
  net_sales numeric,
  total_sales numeric,
  discounts numeric,
  returns_amount numeric,
  refunds_count integer,
  line_items_count numeric,
  financial_status text,
  fulfillment_status text,
  ingested_at_utc timestamptz not null default now()
);

create index if not exists shopify_orders_created_at_idx on public.shopify_orders (created_at_utc);
create index if not exists shopify_orders_customer_type_idx on public.shopify_orders (customer_type);
create index if not exists shopify_orders_source_name_idx on public.shopify_orders (source_name);

create table if not exists public.shopify_order_lines (
  order_line_key text primary key,
  order_id text not null,
  line_item_id text not null,
  created_at_utc timestamptz not null,
  product_id text,
  variant_id text,
  sku text,
  product_title text,
  variant_title text,
  vendor text,
  source_name text,
  customer_type text,
  quantity numeric,
  gross_revenue numeric,
  discount_amount numeric,
  net_revenue numeric,
  returned_quantity numeric,
  returned_revenue numeric,
  net_quantity numeric,
  net_revenue_after_returns numeric,
  ingested_at_utc timestamptz not null default now()
);

create index if not exists shopify_order_lines_created_at_idx on public.shopify_order_lines (created_at_utc);
create index if not exists shopify_order_lines_product_id_idx on public.shopify_order_lines (product_id);
create index if not exists shopify_order_lines_product_title_idx on public.shopify_order_lines (product_title);

alter table public.hourly_metrics enable row level security;
alter table public.shopify_orders enable row level security;
alter table public.shopify_order_lines enable row level security;

drop policy if exists "Allow anon read hourly metrics" on public.hourly_metrics;
create policy "Allow anon read hourly metrics"
on public.hourly_metrics
for select
to anon
using (true);

drop policy if exists "Allow anon read shopify orders" on public.shopify_orders;
create policy "Allow anon read shopify orders"
on public.shopify_orders
for select
to anon
using (true);

drop policy if exists "Allow anon read shopify order lines" on public.shopify_order_lines;
create policy "Allow anon read shopify order lines"
on public.shopify_order_lines
for select
to anon
using (true);
