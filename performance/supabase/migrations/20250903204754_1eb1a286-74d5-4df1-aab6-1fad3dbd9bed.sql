-- Remove the pdf_leads table and all related policies completely
-- This eliminates all data storage and privacy concerns

-- Drop all RLS policies first
DROP POLICY IF EXISTS "Secure lead access policy" ON public.pdf_leads;
DROP POLICY IF EXISTS "Anyone can insert leads (user_id enforced)" ON public.pdf_leads;
DROP POLICY IF EXISTS "Users can delete their own leads" ON public.pdf_leads;
DROP POLICY IF EXISTS "Users can update their own leads" ON public.pdf_leads;

-- Drop the table completely
DROP TABLE IF EXISTS public.pdf_leads;