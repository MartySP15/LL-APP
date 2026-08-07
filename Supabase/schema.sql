-- Run this once in your Supabase project's SQL Editor (Database > SQL Editor > New query)

create extension if not exists pgcrypto;

-- Auto-incrementing human-friendly numbers (ORD-0001, INV-0001, ...)
create sequence if not exists order_number_seq start 1;
create sequence if not exists invoice_number_seq start 1;

create table if not exists customers (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  email text,
  phone text,
  address text,
  created_at timestamptz default now()
);

create table if not exists orders (
  id uuid primary key default gen_random_uuid(),
  order_number text,
  invoice_number text,
  customer_id uuid references customers(id) on delete set null,
  date date,
  due_date date,
  tax_rate numeric default 0,
  notes text,
  items jsonb default '[]'::jsonb,
  order_status text default 'draft',
  amount_paid numeric default 0,
  created_at timestamptz default now()
);

create table if not exists business_settings (
  id int primary key default 1,
  name text default 'Liveable Layouts',
  tagline text default 'Designed to work in real life',
  address text,
  email text,
  payment_terms text default 'Due on receipt',
  payment_method text,
  vat_note text,
  signature_name text,
  social_links text,
  accent text default '#2C6CB0',
  constraint single_row check (id = 1)
);
insert into business_settings (id) values (1) on conflict (id) do nothing;

-- Assign ORD-0001, ORD-0002, ... automatically on insert
create or replace function set_order_number() returns trigger as $$
begin
  if new.order_number is null then
    new.order_number := 'ORD-' || lpad(nextval('order_number_seq')::text, 4, '0');
  end if;
  return new;
end;
$$ language plpgsql;

drop trigger if exists trg_set_order_number on orders;
create trigger trg_set_order_number
before insert on orders
for each row execute function set_order_number();

-- Called from the app to assign INV-0001, INV-0002, ... when an invoice is generated
create or replace function next_invoice_number() returns text as $$
begin
  return 'INV-' || lpad(nextval('invoice_number_seq')::text, 4, '0');
end;
$$ language plpgsql;

-- Row Level Security
-- These policies allow full read/write to anyone with your project's anon key,
-- which is fine for a small trusted team sharing one link. If you later want
-- per-user logins, add Supabase Auth and tighten these policies to check auth.uid().
alter table customers enable row level security;
alter table orders enable row level security;
alter table business_settings enable row level security;

drop policy if exists "allow all customers" on customers;
create policy "allow all customers" on customers for all using (true) with check (true);

drop policy if exists "allow all orders" on orders;
create policy "allow all orders" on orders for all using (true) with check (true);

drop policy if exists "allow all business_settings" on business_settings;
create policy "allow all business_settings" on business_settings for all using (true) with check (true);
