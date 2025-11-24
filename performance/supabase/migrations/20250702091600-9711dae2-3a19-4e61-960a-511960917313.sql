-- Make peso_kg nullable to support running mode (no weight required)
ALTER TABLE public.pdf_leads ALTER COLUMN peso_kg DROP NOT NULL;