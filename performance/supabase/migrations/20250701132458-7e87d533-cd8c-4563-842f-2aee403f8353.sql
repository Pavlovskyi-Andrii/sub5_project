
-- Creare tabella per raccogliere i lead dai PDF reports
CREATE TABLE public.pdf_leads (
  id UUID NOT NULL DEFAULT gen_random_uuid() PRIMARY KEY,
  nome TEXT NOT NULL,
  cognome TEXT NOT NULL,
  email TEXT NOT NULL,
  peso_kg DECIMAL(5,2) NOT NULL CHECK (peso_kg > 0 AND peso_kg <= 300),
  data_generazione TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT now(),
  parametri_fisiologici JSONB NOT NULL,
  created_at TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT now()
);

-- Abilitare Row Level Security
ALTER TABLE public.pdf_leads ENABLE ROW LEVEL SECURITY;

-- Policy per permettere inserimenti pubblici (per la generazione PDF)
CREATE POLICY "Chiunque può creare lead PDF" 
  ON public.pdf_leads 
  FOR INSERT 
  WITH CHECK (true);

-- Policy per lettura solo per utenti autenticati (per eventuali dashboard admin)
CREATE POLICY "Solo utenti autenticati possono leggere lead" 
  ON public.pdf_leads 
  FOR SELECT 
  USING (auth.role() = 'authenticated');

-- Creare indice per ottimizzare ricerche per email e data
CREATE INDEX idx_pdf_leads_email ON public.pdf_leads(email);
CREATE INDEX idx_pdf_leads_data ON public.pdf_leads(data_generazione DESC);
