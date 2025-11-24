-- Fix security vulnerability: Add explicit policy to deny access to anonymous leads
-- This prevents potential exposure of anonymous customer data (names, emails, physiological parameters)

-- Drop the existing SELECT policy to recreate it with proper anonymous lead protection
DROP POLICY IF EXISTS "Users can view their own leads" ON public.pdf_leads;

-- Create a comprehensive SELECT policy that:
-- 1. Allows authenticated users to view their own leads (user_id matches auth.uid())
-- 2. Explicitly denies access to anonymous leads (user_id IS NULL) for all regular users
-- 3. Only allows service role access to anonymous leads for business operations
CREATE POLICY "Secure lead access policy" 
ON public.pdf_leads 
FOR SELECT 
USING (
  -- Allow authenticated users to view their own leads
  (user_id IS NOT NULL AND auth.uid() = user_id)
  -- Anonymous leads (user_id IS NULL) are explicitly denied access
  -- They can only be accessed by service role for business operations
);

-- Add a comment to document the security measure
COMMENT ON POLICY "Secure lead access policy" ON public.pdf_leads IS 
'Allows users to view only their own authenticated leads. Anonymous leads (user_id IS NULL) are protected from unauthorized access and can only be accessed by service role for business operations.';