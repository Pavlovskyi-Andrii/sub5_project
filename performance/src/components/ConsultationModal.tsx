
import React, { useState } from 'react';
import { Button } from '@/components/ui/button';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog';
import { Card, CardContent } from '@/components/ui/card';
import { Mail, Download, ExternalLink } from 'lucide-react';
import { PowerPrecisionModels } from '@/utils/powerPrecisionCalculations';
import PdfLeadForm from './PdfLeadForm';
import { useTranslation } from 'react-i18next';

interface ConsultationModalProps {
  models: PowerPrecisionModels;
  isOpen: boolean;
  onOpenChange: (open: boolean) => void;
  isRunning?: boolean; // Add isRunning prop
  weight?: number; // Add weight prop
  gender?: 'male' | 'female'; // Add gender prop
}

const ConsultationModal: React.FC<ConsultationModalProps> = ({ 
  models, 
  isOpen, 
  onOpenChange, 
  isRunning = false,
  weight,
  gender
}) => {
  const { t } = useTranslation();
  
  const mailtoLink = `mailto:info@miccoach.it?subject=${encodeURIComponent('Richiesta Consulenza Performance - Report da MicCoach Lab')}&body=${encodeURIComponent(`Ciao Michael,

Ti contatto tramite l'app MicCoach Lab. Vorrei richiedere una consulenza per analizzare il mio profilo di performance e discutere delle mie esigenze di allenamento.

In allegato trovi il mio report generato dall'app.

Resto in attesa di un tuo riscontro.

Saluti,
[Inserisci qui il tuo nome]`)}`;

  return (
    <Dialog open={isOpen} onOpenChange={onOpenChange}>
      <DialogTrigger asChild>
        <Button 
          variant="outline"
          className="border-lime-400 text-lime-400 hover:bg-lime-400 hover:text-slate-900 font-semibold transition-all duration-200 hover:scale-105"
          disabled={!models.valid}
        >
          <Mail className="mr-2 h-4 w-4" />
          {t('consultation.trigger_button')}
        </Button>
      </DialogTrigger>
      <DialogContent className="sm:max-w-lg bg-slate-800 border-slate-700">
        <DialogHeader>
          <DialogTitle className="text-lime-400 text-xl">{t('consultation.title')}</DialogTitle>
        </DialogHeader>
        <Card className="bg-slate-900 border-slate-700">
          <CardContent className="p-6 space-y-6">
            <p className="text-slate-200 text-center">
              {t('consultation.subtitle')}
            </p>
            
            <div className="space-y-4">
              <div className="bg-slate-800 p-4 rounded-lg border border-slate-600">
                <div className="flex items-center mb-3">
                  <div className="bg-lime-400 text-slate-900 rounded-full w-6 h-6 flex items-center justify-center text-sm font-bold mr-3">
                    1
                  </div>
                  <h3 className="text-slate-200 font-semibold">{t('consultation.step1.title')}</h3>
                </div>
                <p className="text-slate-400 text-sm mb-3">
                  {t('consultation.step1.description')}
                </p>
                <PdfLeadForm models={models} isRunning={isRunning} weight={weight} gender={gender} />
              </div>

              <div className="bg-slate-800 p-4 rounded-lg border border-slate-600">
                <div className="flex items-center mb-3">
                  <div className="bg-lime-400 text-slate-900 rounded-full w-6 h-6 flex items-center justify-center text-sm font-bold mr-3">
                    2
                  </div>
                  <h3 className="text-slate-200 font-semibold">{t('consultation.step2.title')}</h3>
                </div>
                <p className="text-slate-400 text-sm mb-3">
                  {t('consultation.step2.description')}
                </p>
                <Button 
                  className="w-full bg-slate-700 hover:bg-slate-600 text-lime-400 border border-lime-400"
                  onClick={() => window.open(mailtoLink, '_blank')}
                >
                  <ExternalLink className="mr-2 h-4 w-4" />
                  {t('consultation.step2.button')}
                </Button>
              </div>
            </div>

            <div className="text-center pt-4 border-t border-slate-700">
              <p className="text-lime-400 font-semibold">
                {t('consultation.footer')}
              </p>
            </div>
          </CardContent>
        </Card>
      </DialogContent>
    </Dialog>
  );
};

export default ConsultationModal;
