
import React, { useRef } from 'react';
import { Dialog, DialogContent, DialogHeader, DialogTitle } from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { Download } from 'lucide-react';
import { AthleteProfile } from '@/utils/athleteCardCalculations';
import { RunningAthleteProfile } from '@/utils/runningAthleteCalculations';
import { formatProfileType } from '@/utils/athleteCardCalculations';
import { formatRunningProfileType } from '@/utils/runningAthleteCalculations';
import { Country } from '@/utils/countries';
import InstagramStoryTemplate from './InstagramStoryTemplate';
import html2canvas from 'html2canvas';
import { useToast } from '@/hooks/use-toast';

interface SocialShareModalProps {
  isOpen: boolean;
  onClose: () => void;
  athleteName: string;
  country: Country;
  profileImage: string;
  profile: AthleteProfile | RunningAthleteProfile;
  isRunning?: boolean;
}

const SocialShareModal: React.FC<SocialShareModalProps> = ({
  isOpen,
  onClose,
  athleteName,
  country,
  profileImage,
  profile,
  isRunning = false
}) => {
  const storyRef = useRef<HTMLDivElement>(null);
  const { toast } = useToast();

  const profileTypeName = isRunning 
    ? formatRunningProfileType(profile.profileType)
    : formatProfileType(profile.profileType);

  const downloadCard = async () => {
    if (!storyRef.current) return;

    try {
      // Allow time for rendering
      await new Promise(resolve => setTimeout(resolve, 1000));

      const canvas = await html2canvas(storyRef.current, {
        scale: 2,
        useCORS: true,
        backgroundColor: null,
        logging: false,
        allowTaint: true,
        width: 1080,
        height: 1920,
      });

      canvas.toBlob((blob) => {
        if (blob) {
          const url = URL.createObjectURL(blob);
          const link = document.createElement('a');
          link.href = url;
          const prefix = isRunning ? 'running-' : '';
          link.download = `miccoach-${prefix}card-${athleteName.replace(/\s+/g, '-').toLowerCase()}.png`;
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          URL.revokeObjectURL(url);

          toast({
            title: "Card scaricata! 🎉",
            description: "La tua card è stata scaricata con successo.",
          });
        }
      }, 'image/png', 0.9);
    } catch (error) {
      console.error('Errore nel download:', error);
      toast({
        title: "Errore",
        description: "Si è verificato un errore durante il download.",
        variant: "destructive"
      });
    }
  };

  return (
    <>
      <Dialog open={isOpen} onOpenChange={onClose}>
        <DialogContent className="sm:max-w-md bg-slate-800 border-slate-700">
          <DialogHeader>
            <DialogTitle className="text-lime-400 text-2xl text-center">
              Complimenti, {athleteName}! 🎉
            </DialogTitle>
          </DialogHeader>

          <div className="space-y-6 text-center">
            {/* Testo di incoraggiamento */}
            <div>
              <p className="text-slate-200 text-lg">
                La tua Card Atleta è pronta!
              </p>
              <p className="text-slate-400 mt-1">
                Mostra al mondo il tuo profilo da <span className="text-lime-400 font-semibold">{profileTypeName}</span>
              </p>
            </div>

            {/* Pulsante di download principale */}
            <Button
              onClick={downloadCard}
              className="w-full bg-gradient-to-r from-lime-400 to-green-500 hover:from-lime-500 hover:to-green-600 text-slate-900 font-semibold text-lg py-6"
            >
              <Download className="mr-3 h-6 w-6" />
              Scarica la Card
            </Button>

            {/* Istruzioni per la condivisione */}
            <div className="bg-slate-900 p-4 rounded-lg border border-slate-700">
              <p className="text-slate-200 font-semibold mb-2">
                Pronto a mostrarla al mondo?
              </p>
              <p className="text-slate-300 text-sm">
                Taggaci su Instagram <span className="text-lime-400 font-semibold">@MicCoach</span> e usa{' '}
                <span className="text-lime-400 font-semibold">#MicCoachLab</span> per apparire nella nostra gallery!
              </p>
            </div>
          </div>
        </DialogContent>
      </Dialog>

      {/* Hidden element for download */}
      <div className="fixed -left-[9999px] -top-[9999px] pointer-events-none" style={{ zIndex: -1 }}>
        <div ref={storyRef}>
          <InstagramStoryTemplate
            athleteName={athleteName}
            country={country}
            profileImage={profileImage}
            profile={profile}
            isRunning={isRunning}
          />
        </div>
      </div>
    </>
  );
};

export default SocialShareModal;
