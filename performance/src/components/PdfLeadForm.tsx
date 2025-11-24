import React, { useState } from 'react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Checkbox } from '@/components/ui/checkbox';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog';
import { Card, CardContent } from '@/components/ui/card';
import { Download, Loader2, Upload, X } from 'lucide-react';
import { PowerPrecisionModels } from '@/utils/powerPrecisionCalculations';
import { PrecisionModels } from '@/utils/precisionCalculations';
import { useToast } from '@/hooks/use-toast';
import { generatePdfReport } from './PdfReportGenerator';
import { generateRunningPdfReport } from './RunningPdfReportGenerator';
import { PDF_COLOR_PALETTES, getPalettePreviewColors } from '@/utils/pdfColorPalettes';
import html2canvas from 'html2canvas';

interface PdfLeadFormProps {
  models: PowerPrecisionModels | PrecisionModels;
  onPdfGenerated?: () => void;
  isRunning?: boolean;
  weight?: number; // Weight from parent component for calculations
  gender?: 'male' | 'female'; // Gender from parent component
}

interface CoachData {
  isCoach: boolean;
  coachName?: string;
  coachCompany?: string;
  coachEmail?: string;
  coachLogo?: string;
}

interface FormData {
  nome: string;
  cognome: string;
  consenso: boolean;
  isCoach: boolean;
  coachName: string;
  coachCompany: string;
  coachEmail: string;
  coachLogo: string;
  colorPalette: string;
}

const PdfLeadForm: React.FC<PdfLeadFormProps> = ({ models, onPdfGenerated, isRunning = false, weight, gender }) => {
  const [formData, setFormData] = useState<FormData>({
    nome: '',
    cognome: '',
    consenso: false,
    isCoach: false,
    coachName: '',
    coachCompany: '',
    coachEmail: '',
    coachLogo: '',
    colorPalette: 'green'
  });
  const [isLoading, setIsLoading] = useState(false);
  const [errors, setErrors] = useState<Partial<Record<keyof FormData, string>>>({});
  const [isOpen, setIsOpen] = useState(false);
  const { toast } = useToast();

  const validateForm = (): boolean => {
    const newErrors: Partial<Record<keyof FormData, string>> = {};
    
    if (!formData.nome.trim()) newErrors.nome = 'Il nome è obbligatorio';
    if (!formData.cognome.trim()) newErrors.cognome = 'Il cognome è obbligatorio';
    
    if (formData.isCoach) {
      if (!formData.coachName.trim()) newErrors.coachName = 'Il nome del coach è obbligatorio';
      if (!formData.coachEmail.trim()) newErrors.coachEmail = 'L\'email del coach è obbligatoria';
      else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(formData.coachEmail)) {
        newErrors.coachEmail = 'Inserisci un\'email valida';
      }
    }
    
    if (!formData.consenso) {
      newErrors.consenso = 'Devi accettare il consenso per procedere';
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const captureChartImage = async (): Promise<string | null> => {
    try {
      const chartElementId = isRunning ? 'running-chart-container' : 'power-chart-container';
      const chartElement = document.getElementById(chartElementId);
      if (!chartElement) {
        console.warn('Elemento del grafico non trovato');
        return null;
      }

      await new Promise(resolve => setTimeout(resolve, 1000));

      const canvas = await html2canvas(chartElement, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#1e293b',
        width: chartElement.offsetWidth,
        height: chartElement.offsetHeight
      });

      return canvas.toDataURL('image/png');
    } catch (error) {
      console.error('Errore nella cattura del grafico:', error);
      return null;
    }
  };

  const captureZonesImage = async (): Promise<string | null> => {
    try {
      const zonesElementId = isRunning ? 'running-zones-container' : 'power-zones-container';
      const zonesElement = document.getElementById(zonesElementId);
      if (!zonesElement) {
        console.warn('Elemento delle zone non trovato');
        return null;
      }

      await new Promise(resolve => setTimeout(resolve, 1000));

      const canvas = await html2canvas(zonesElement, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#1e293b',
        width: zonesElement.offsetWidth,
        height: zonesElement.offsetHeight
      });

      return canvas.toDataURL('image/png');
    } catch (error) {
      console.error('Errore nella cattura delle zone:', error);
      return null;
    }
  };

  const captureProfileImage = async (): Promise<string | null> => {
    try {
      const profileElementId = isRunning ? 'running-physiological-profile' : 'physiological-profile';
      const profileElement = document.getElementById(profileElementId);
      if (!profileElement) {
        console.warn('Elemento del profilo fisiologico non trovato');
        return null;
      }

      await new Promise(resolve => setTimeout(resolve, 1000));

      const canvas = await html2canvas(profileElement, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#1e293b',
        width: profileElement.offsetWidth,
        height: profileElement.offsetHeight
      });

      return canvas.toDataURL('image/png');
    } catch (error) {
      console.error('Errore nella cattura del profilo:', error);
      return null;
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    
    if (!validateForm()) return;
    if (!models.valid) {
      toast({
        title: "Errore",
        description: "I modelli di calcolo non sono validi. Riprova dopo aver inserito dei dati validi.",
        variant: "destructive"
      });
      return;
    }

    setIsLoading(true);
    try {
      console.log('Inizio generazione PDF client-side con dati:', { formData, models, isRunning, weight, gender });
      
      const chartImageData = await captureChartImage();
      const zonesImageData = await captureZonesImage();
      const profileImageData = await captureProfileImage();
      
      // Prepara i dati utente e coach completamente client-side
      const userData = {
        nome: formData.nome,
        cognome: formData.cognome,
        email: `${formData.nome.toLowerCase()}.${formData.cognome.toLowerCase()}@report.local`, // Email fittizia per PDF
        peso_kg: weight || 70 // Sempre includi il peso (default 70kg)
      };

      const coachData: CoachData = {
        isCoach: formData.isCoach,
        coachName: formData.coachName || undefined,
        coachCompany: formData.coachCompany || undefined,
        coachEmail: formData.coachEmail || undefined,
        coachLogo: formData.coachLogo || undefined
      };

      // Genera il PDF completamente lato client (senza database)
      console.log('Generazione PDF completamente client-side...');
      let pdfBlob;
      
      if (isRunning) {
        pdfBlob = await generateRunningPdfReport(userData, models as PrecisionModels, chartImageData, zonesImageData, profileImageData, coachData, formData.colorPalette);
      } else {
        pdfBlob = await generatePdfReport(userData, models as PowerPrecisionModels, chartImageData, profileImageData, coachData, formData.colorPalette);
      }

      // Scarica il PDF
      const url = window.URL.createObjectURL(pdfBlob);
      const link = document.createElement('a');
      link.href = url;
      const sportType = isRunning ? 'Running' : 'Ciclismo';
      link.download = `Report_${sportType}_${formData.nome}_${formData.cognome}.pdf`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      // Reset form e chiudi dialog
    setFormData({
      nome: '',
      cognome: '',
      consenso: false,
      isCoach: false,
      coachName: '',
      coachCompany: '',
      coachEmail: '',
      coachLogo: '',
      colorPalette: 'green'
    });
      setIsOpen(false);
      onPdfGenerated?.();
      
      toast({
        title: "Successo",
        description: "Report PDF generato e scaricato con successo! 🔒 Nessun dato salvato sui nostri server.",
        variant: "default"
      });
      
    } catch (error) {
      console.error('Errore completo:', error);
      const errorMessage = error instanceof Error ? error.message : 'Errore sconosciuto nella generazione del PDF';
      
      toast({
        title: "Errore nella generazione del PDF",
        description: errorMessage + ". Riprova più tardi.",
        variant: "destructive"
      });
    } finally {
      setIsLoading(false);
    }
  };

  const handleInputChange = (field: keyof Omit<FormData, 'consenso' | 'isCoach'>, value: string) => {
    setFormData(prev => ({ ...prev, [field]: value }));
    if (errors[field]) {
      setErrors(prev => ({ ...prev, [field]: undefined }));
    }
  };

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 2 * 1024 * 1024) { // 2MB limit
        setErrors(prev => ({ ...prev, coachLogo: 'Il file deve essere massimo 2MB' }));
        return;
      }
      
      const reader = new FileReader();
      reader.onload = (event) => {
        const base64 = event.target?.result as string;
        setFormData(prev => ({ ...prev, coachLogo: base64 }));
        setErrors(prev => ({ ...prev, coachLogo: undefined }));
      };
      reader.readAsDataURL(file);
    }
  };

  const handleCheckboxChange = (field: 'consenso' | 'isCoach') => (checked: boolean | 'indeterminate') => {
    const booleanValue = checked === true;
    setFormData(prev => ({ ...prev, [field]: booleanValue }));
    if (errors[field]) {
      setErrors(prev => ({ ...prev, [field]: undefined }));
    }
    
    // Clear coach fields if not a coach
    if (field === 'isCoach' && !booleanValue) {
      setFormData(prev => ({ 
        ...prev, 
        isCoach: false,
        coachName: '',
        coachCompany: '',
        coachEmail: '',
        coachLogo: ''
      }));
    }
  };

  return (
    <Dialog open={isOpen} onOpenChange={setIsOpen}>
      <DialogTrigger asChild>
        <Button 
          className="bg-lime-400 hover:bg-lime-500 text-slate-900 font-semibold transition-all duration-200 hover:scale-105 hover:shadow-lg"
          disabled={!models.valid}
        >
          <Download className="mr-2 h-4 w-4" />
          Scarica la tua Analisi Completa
        </Button>
      </DialogTrigger>
      <DialogContent className="sm:max-w-md bg-slate-800 border-slate-700">
        <DialogHeader>
          <DialogTitle className="text-lime-400">
            Genera Report {isRunning ? 'Running' : 'Prestazioni'}
          </DialogTitle>
          <p className="text-sm text-slate-400">
            🔒 Privacy garantita: i tuoi dati non vengono salvati sui nostri server
          </p>
        </DialogHeader>
        <Card className="bg-slate-900 border-slate-700">
          <CardContent className="p-6">
            <form onSubmit={handleSubmit} className="space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label className="text-slate-200">Nome *</Label>
                  <Input
                    value={formData.nome}
                    onChange={(e) => handleInputChange('nome', e.target.value)}
                    className="bg-slate-700 border-slate-600 text-white"
                    placeholder="Inserisci il nome"
                  />
                  {errors.nome && <p className="text-red-400 text-sm mt-1">{errors.nome}</p>}
                </div>
                <div>
                  <Label className="text-slate-200">Cognome *</Label>
                  <Input
                    value={formData.cognome}
                    onChange={(e) => handleInputChange('cognome', e.target.value)}
                    className="bg-slate-700 border-slate-600 text-white"
                    placeholder="Inserisci il cognome"
                  />
                  {errors.cognome && <p className="text-red-400 text-sm mt-1">{errors.cognome}</p>}
                </div>
              </div>
              
              
              {!isRunning && weight && (
                <div className="bg-slate-800 p-3 rounded-md">
                  <p className="text-slate-300 text-sm">
                    <strong>Peso utilizzato per i calcoli:</strong> {weight} kg
                  </p>
                  <p className="text-slate-400 text-xs mt-1">
                    Il peso viene utilizzato solo per calcolare i tuoi W/kg e non viene salvato
                  </p>
                </div>
              )}

              <div className="space-y-6">
                <div className="flex items-center space-x-4 p-4 bg-slate-700 rounded-lg border border-slate-600 hover:bg-slate-650 transition-colors">
                  <Checkbox
                    id="isCoach"
                    checked={formData.isCoach}
                    onCheckedChange={handleCheckboxChange('isCoach')}
                    className="w-6 h-6 border-2 border-lime-400 data-[state=checked]:bg-lime-400 data-[state=checked]:border-lime-400"
                  />
                  <Label htmlFor="isCoach" className="text-slate-200 text-base font-medium cursor-pointer">
                    💼 Sei un coach/allenatore? Personalizza il report con i tuoi dati
                  </Label>
                </div>

                {formData.isCoach && (
                  <div className="space-y-3 bg-slate-800 p-4 rounded-md border border-slate-600">
                    <h4 className="text-lime-400 font-semibold text-sm">Dati Coach</h4>
                    
                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <Label className="text-slate-200 text-sm">Nome/Ragione Sociale *</Label>
                        <Input
                          value={formData.coachName}
                          onChange={(e) => handleInputChange('coachName', e.target.value)}
                          className="bg-slate-700 border-slate-600 text-white text-sm"
                          placeholder="Es. Mario Rossi"
                        />
                        {errors.coachName && <p className="text-red-400 text-xs mt-1">{errors.coachName}</p>}
                      </div>
                      
                      <div>
                        <Label className="text-slate-200 text-sm">Azienda/Società</Label>
                        <Input
                          value={formData.coachCompany}
                          onChange={(e) => handleInputChange('coachCompany', e.target.value)}
                          className="bg-slate-700 border-slate-600 text-white text-sm"
                          placeholder="Es. Team Rossi Coaching"
                        />
                      </div>
                    </div>

                    <div>
                      <Label className="text-slate-200 text-sm">Email Coach *</Label>
                      <Input
                        type="email"
                        value={formData.coachEmail}
                        onChange={(e) => handleInputChange('coachEmail', e.target.value)}
                        className="bg-slate-700 border-slate-600 text-white text-sm"
                        placeholder="coach@esempio.it"
                      />
                      {errors.coachEmail && <p className="text-red-400 text-xs mt-1">{errors.coachEmail}</p>}
                    </div>

                    <div>
                      <Label className="text-slate-200 text-sm">Logo (opzionale)</Label>
                      <div className="flex items-center gap-3">
                        <label className="cursor-pointer">
                          <div className="flex items-center gap-2 bg-slate-700 hover:bg-slate-600 text-white px-3 py-2 rounded-md border border-slate-600 text-sm transition-colors">
                            <Upload className="h-4 w-4" />
                            Carica Logo
                          </div>
                          <input
                            type="file"
                            accept="image/*"
                            onChange={handleFileChange}
                            className="hidden"
                          />
                        </label>
                        
                        {formData.coachLogo && (
                          <div className="flex items-center gap-2">
                            <img 
                              src={formData.coachLogo} 
                              alt="Logo preview" 
                              className="w-8 h-8 object-contain rounded border border-slate-600" 
                            />
                            <Button
                              type="button"
                              variant="ghost"
                              size="sm"
                              onClick={() => setFormData(prev => ({ ...prev, coachLogo: '' }))}
                              className="text-red-400 hover:text-red-300 hover:bg-red-900/20 p-1 h-auto"
                            >
                              <X className="h-3 w-3" />
                            </Button>
                          </div>
                        )}
                      </div>
                      {errors.coachLogo && <p className="text-red-400 text-xs mt-1">{errors.coachLogo}</p>}
                      <p className="text-slate-400 text-xs mt-1">Formato supportato: JPG, PNG. Max 2MB</p>
                    </div>
                  </div>
                )}
              </div>

              <div className="space-y-4">
                <Label className="text-base font-medium text-slate-200">Scegli la palette di colori del report</Label>
                <div className="grid grid-cols-2 md:grid-cols-3 gap-3">
                  {Object.entries(PDF_COLOR_PALETTES).map(([key, palette]) => (
                    <div
                      key={key}
                      className={`relative cursor-pointer p-3 rounded-lg border-2 transition-all ${
                        formData.colorPalette === key 
                          ? 'border-lime-400 ring-2 ring-lime-400/20' 
                          : 'border-slate-600 hover:border-lime-400/50'
                      }`}
                      onClick={() => setFormData({ ...formData, colorPalette: key })}
                    >
                      <div className="flex items-center gap-2 mb-2">
                        <div 
                          className="w-4 h-4 rounded-full" 
                          style={{ backgroundColor: palette.primary }}
                        />
                        <span className="text-sm font-medium text-slate-200">{palette.name}</span>
                      </div>
                      <div className="flex gap-1">
                        <div 
                          className="w-3 h-3 rounded" 
                          style={{ backgroundColor: palette.primary }}
                        />
                        <div 
                          className="w-3 h-3 rounded" 
                          style={{ backgroundColor: palette.accent }}
                        />
                        <div 
                          className="w-3 h-3 rounded" 
                          style={{ backgroundColor: palette.background }}
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="flex items-center space-x-4 p-4 bg-gradient-to-r from-lime-900/20 to-lime-800/20 rounded-lg border-2 border-lime-400/50 hover:border-lime-400 transition-all">
                <Checkbox
                  id="consenso"
                  checked={formData.consenso}
                  onCheckedChange={handleCheckboxChange('consenso')}
                  className="w-6 h-6 border-2 border-lime-400 data-[state=checked]:bg-lime-400 data-[state=checked]:border-lime-400"
                />
                <Label htmlFor="consenso" className="text-slate-200 text-base font-medium cursor-pointer">
                  ✅ Confermo di voler generare il report PDF con i miei dati (processati solo localmente) *
                </Label>
              </div>
              {errors.consenso && <p className="text-red-400 text-sm">{errors.consenso}</p>}

              <Button 
                type="submit" 
                className="w-full bg-lime-400 hover:bg-lime-500 text-slate-900 font-semibold"
                disabled={isLoading}
              >
                {isLoading ? (
                  <>
                    <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                    Generando PDF...
                  </>
                ) : (
                  <>
                    <Download className="mr-2 h-4 w-4" />
                    Genera e Scarica PDF
                  </>
                )}
              </Button>
            </form>
          </CardContent>
        </Card>
      </DialogContent>
    </Dialog>
  );
};

export default PdfLeadForm;
