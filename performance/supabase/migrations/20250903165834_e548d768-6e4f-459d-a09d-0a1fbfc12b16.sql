-- Add user_id column to pdf_leads table to associate records with users
ALTER TABLE public.pdf_leads 
ADD COLUMN user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE;

-- Set user_id to be non-nullable for future records (existing records will remain null temporarily)
-- We'll handle existing data separately if needed
ALTER TABLE public.pdf_leads 
ALTER COLUMN user_id SET NOT NULL;

-- Drop the existing overly permissive policy
DROP POLICY "Solo utenti autenticati possono leggere lead" ON public.pdf_leads;

-- Create secure user-specific policies
CREATE POLICY "Users can only view their own PDF leads" 
ON public.pdf_leads 
FOR SELECT 
USING (auth.uid() = user_id);

-- Update the insert policy to automatically set user_id
DROP POLICY "Chiunque può creare lead PDF" ON public.pdf_leads;

CREATE POLICY "Users can create their own PDF leads" 
ON public.pdf_leads 
FOR INSERT 
WITH CHECK (auth.uid() = user_id);

-- Add policy for users to update their own records (if needed)
CREATE POLICY "Users can update their own PDF leads" 
ON public.pdf_leads 
FOR UPDATE 
USING (auth.uid() = user_id) 
WITH CHECK (auth.uid() = user_id);

-- Add policy for users to delete their own records (if needed)
CREATE POLICY "Users can delete their own PDF leads" 
ON public.pdf_leads 
FOR DELETE 
USING (auth.uid() = user_id);