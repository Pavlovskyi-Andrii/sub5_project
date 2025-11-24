-- 1) Add nullable user_id and index (do not break existing inserts)
ALTER TABLE public.pdf_leads
  ADD COLUMN IF NOT EXISTS user_id uuid;

-- Add FK to auth.users (safe, sets NULL on user deletion)
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_constraint
    WHERE conname = 'pdf_leads_user_id_fkey'
  ) THEN
    ALTER TABLE public.pdf_leads
      ADD CONSTRAINT pdf_leads_user_id_fkey
      FOREIGN KEY (user_id) REFERENCES auth.users(id)
      ON DELETE SET NULL;
  END IF;
END $$;

CREATE INDEX IF NOT EXISTS idx_pdf_leads_user_id ON public.pdf_leads(user_id);

-- 2) Tighten RLS policies: only owner can SELECT; allow anonymous INSERT with NULL user_id
DROP POLICY IF EXISTS "Solo utenti autenticati possono leggere lead" ON public.pdf_leads;
DROP POLICY IF EXISTS "Chiunque può creare lead PDF" ON public.pdf_leads;
DROP POLICY IF EXISTS "Users can view their own leads" ON public.pdf_leads;
DROP POLICY IF EXISTS "Anyone can insert leads (user_id enforced)" ON public.pdf_leads;
DROP POLICY IF EXISTS "Users can update their own leads" ON public.pdf_leads;
DROP POLICY IF EXISTS "Users can delete their own leads" ON public.pdf_leads;

CREATE POLICY "Users can view their own leads"
ON public.pdf_leads
FOR SELECT
USING (user_id IS NOT NULL AND auth.uid() = user_id);

CREATE POLICY "Anyone can insert leads (user_id enforced)"
ON public.pdf_leads
FOR INSERT
WITH CHECK (user_id IS NULL OR auth.uid() = user_id);

CREATE POLICY "Users can update their own leads"
ON public.pdf_leads
FOR UPDATE
USING (user_id IS NOT NULL AND auth.uid() = user_id)
WITH CHECK (user_id IS NOT NULL AND auth.uid() = user_id);

CREATE POLICY "Users can delete their own leads"
ON public.pdf_leads
FOR DELETE
USING (user_id IS NOT NULL AND auth.uid() = user_id);