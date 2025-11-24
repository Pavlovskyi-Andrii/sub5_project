
import React, { useState, useRef } from 'react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog';
import { Upload, Image as ImageIcon } from 'lucide-react';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import { PowerPrecisionModels } from '@/utils/powerPrecisionCalculations';
import { calculateAthleteProfile, Gender } from '@/utils/athleteCardCalculations';
import { COUNTRIES, getCountryByCode, Country } from '@/utils/countries';
import ImageCropper from './ImageCropper';
import InstagramStoryTemplate from './InstagramStoryTemplate';
import SocialShareModal from './SocialShareModal';
import { useToast } from '@/hooks/use-toast';

// Avatar di default
import avatarMaleVector from '@/assets/avatar-male-vector.png';
import avatarMaleManga from '@/assets/avatar-male-manga.png';
import avatarMaleShonen from '@/assets/avatar-male-shonen.png';
import avatarFemaleVector from '@/assets/avatar-female-vector.png';
import avatarFemaleManga from '@/assets/avatar-female-manga.png';
import avatarFemaleShonen from '@/assets/avatar-female-shonen.png';

interface AthleteCardGeneratorProps {
  models: PowerPrecisionModels;
  bodyWeight: number;
  gender: 'male' | 'female';
}

const getMaleAvatars = () => [
  { id: 'male-vector', src: avatarMaleVector, name: 'Vettoriale' },
  { id: 'male-manga', src: avatarMaleManga, name: 'Manga' },
  { id: 'male-shonen', src: avatarMaleShonen, name: 'Shōnen' }
];

const getFemaleAvatars = () => [
  { id: 'female-vector', src: avatarFemaleVector, name: 'Vettoriale' },
  { id: 'female-manga', src: avatarFemaleManga, name: 'Manga' },
  { id: 'female-shonen', src: avatarFemaleShonen, name: 'Shōnen' }
];

const AthleteCardGenerator: React.FC<AthleteCardGeneratorProps> = ({
  models,
  bodyWeight,
  gender: initialGender
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [athleteName, setAthleteName] = useState('');
  const [gender, setGender] = useState<Gender>(initialGender);
  const [selectedCountry, setSelectedCountry] = useState<Country | null>(null);
  const [profileImage, setProfileImage] = useState<string>('');
  const [selectedAvatar, setSelectedAvatar] = useState<string>('');
  const [showImageCropper, setShowImageCropper] = useState(false);
  const [tempImageSrc, setTempImageSrc] = useState<string>('');
  const [showSocialModal, setShowSocialModal] = useState(false);
  
  const fileInputRef = useRef<HTMLInputElement>(null);
  const { toast } = useToast();

  // Calcola il profilo dell'atleta
  const athleteProfile = calculateAthleteProfile(models, bodyWeight, gender);

  // Avatar dinamici basati sul genere
  const availableAvatars = gender === 'male' ? getMaleAvatars() : getFemaleAvatars();

  const handleImageUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const imageSrc = e.target?.result as string;
        setTempImageSrc(imageSrc);
        setShowImageCropper(true);
      };
      reader.readAsDataURL(file);
    }
  };

  const handleGenderChange = (value: Gender) => {
    setGender(value);
    // Reset avatar selection when gender changes
    setSelectedAvatar('');
    setProfileImage('');
  };

  const handleCropComplete = (croppedImageUrl: string) => {
    setProfileImage(croppedImageUrl);
    setSelectedAvatar('');
  };

  const handleAvatarSelect = (avatarSrc: string) => {
    setSelectedAvatar(avatarSrc);
    setProfileImage(avatarSrc);
  };

  const handleCountrySelect = (countryCode: string) => {
    const country = getCountryByCode(countryCode);
    if (country) {
      setSelectedCountry(country);
    }
  };

  const canGenerateCard = () => {
    return athleteName.trim() !== '' && 
           gender !== null &&
           selectedCountry !== null && 
           (profileImage !== '' || selectedAvatar !== '') &&
           models.valid;
  };

  const generateAndShowModal = () => {
    if (!canGenerateCard()) return;
    setShowSocialModal(true);
  };

  const resetForm = () => {
    setAthleteName('');
    setGender(initialGender);
    setSelectedCountry(null);
    setProfileImage('');
    setSelectedAvatar('');
  };

  return (
    <>
      <Dialog open={isOpen} onOpenChange={setIsOpen}>
        <DialogTrigger asChild>
          <Button 
            className="bg-gradient-to-r from-lime-400 to-green-500 hover:from-lime-500 hover:to-green-600 text-slate-900 font-semibold transition-all duration-200 hover:scale-105 hover:shadow-lg"
            disabled={!models.valid}
          >
            <ImageIcon className="mr-2 h-4 w-4" />
            Crea Card Atleta
          </Button>
        </DialogTrigger>
        
        <DialogContent className="sm:max-w-5xl bg-slate-800 border-slate-700 max-h-[80vh] overflow-y-auto">
          <DialogHeader>
            <DialogTitle className="text-lime-400 text-2xl">
              Genera la tua Card Atleta
            </DialogTitle>
          </DialogHeader>

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            {/* Form di input - più stretto */}
            <Card className="bg-slate-900 border-slate-700 lg:col-span-1">
              <CardHeader className="pb-4">
                <CardTitle className="text-lime-400">Personalizza la tua Card</CardTitle>
              </CardHeader>
              <CardContent className="space-y-3">
                {/* Nome Atleta */}
                <div>
                  <Label className="text-slate-200">Nome Atleta</Label>
                  <Input
                    value={athleteName}
                    onChange={(e) => setAthleteName(e.target.value)}
                    className="bg-slate-700 border-slate-600 text-white"
                    placeholder="Inserisci il tuo nome"
                  />
                </div>

                 {/* Sesso */}
                 <div>
                   <Label className="text-slate-200">Sesso</Label>
                   <RadioGroup 
                     value={gender} 
                     onValueChange={handleGenderChange}
                     className="flex space-x-6 mt-2"
                   >
                     <div className="flex items-center space-x-2">
                       <RadioGroupItem value="male" id="male" />
                       <Label 
                         htmlFor="male" 
                         className={`px-3 py-1 rounded cursor-pointer transition-colors ${
                           gender === 'male' 
                             ? 'bg-lime-400 text-slate-900 font-semibold' 
                             : 'text-slate-300 hover:text-white'
                         }`}
                       >
                         Maschile
                       </Label>
                     </div>
                     <div className="flex items-center space-x-2">
                       <RadioGroupItem value="female" id="female" />
                       <Label 
                         htmlFor="female" 
                         className={`px-3 py-1 rounded cursor-pointer transition-colors ${
                           gender === 'female' 
                             ? 'bg-lime-400 text-slate-900 font-semibold' 
                             : 'text-slate-300 hover:text-white'
                         }`}
                       >
                         Femminile
                       </Label>
                     </div>
                   </RadioGroup>
                 </div>

                 {/* Nazionalità */}
                <div>
                  <Label className="text-slate-200">Nazionalità</Label>
                  <Select onValueChange={handleCountrySelect}>
                    <SelectTrigger className="bg-slate-700 border-slate-600 text-white">
                      <SelectValue placeholder="Seleziona la tua nazionalità" />
                    </SelectTrigger>
                    <SelectContent className="bg-slate-700 border-slate-600">
                      {COUNTRIES.map((country) => (
                        <SelectItem 
                          key={country.code} 
                          value={country.code}
                          className="text-white hover:bg-slate-600"
                        >
                          <span className="flex items-center space-x-2">
                            <span>{country.flag}</span>
                            <span>{country.name}</span>
                          </span>
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                {/* Immagine profilo */}
                <div>
                  <Label className="text-slate-200">Immagine Profilo</Label>
                  
                  {/* Upload personalizzato */}
                  <div className="mt-2">
                    <Button
                      onClick={() => fileInputRef.current?.click()}
                      variant="outline"
                      className="w-full border-slate-600 text-slate-200 hover:bg-slate-700"
                    >
                      <Upload className="mr-2 h-4 w-4" />
                      Carica la tua foto
                    </Button>
                    <input
                      ref={fileInputRef}
                      type="file"
                      accept="image/*"
                      onChange={handleImageUpload}
                      className="hidden"
                    />
                  </div>

                  {/* Avatar predefiniti */}
                  <div className="mt-3">
                    <Label className="text-slate-300 text-sm">O scegli un avatar ({gender === 'male' ? 'maschile' : 'femminile'}):</Label>
                    <div className="grid grid-cols-3 gap-2 mt-2">
                      {availableAvatars.map((avatar) => (
                        <button
                          key={avatar.id}
                          onClick={() => handleAvatarSelect(avatar.src)}
                          className={`p-2 rounded-lg border-2 transition-all ${
                            selectedAvatar === avatar.src
                              ? 'border-lime-400 bg-lime-400/10'
                              : 'border-slate-600 hover:border-slate-500'
                          }`}
                        >
                          <img
                            src={avatar.src}
                            alt={avatar.name}
                            className="w-full h-12 object-cover rounded"
                          />
                          <p className="text-xs text-slate-300 mt-1">{avatar.name}</p>
                        </button>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Anteprima immagine selezionata */}
                {profileImage && (
                  <div className="mt-3">
                    <Label className="text-slate-200">Anteprima</Label>
                    <div className="mt-2 w-20 h-20 rounded-full overflow-hidden border-2 border-lime-400">
                      <img
                        src={profileImage}
                        alt="Preview"
                        className="w-full h-full object-cover"
                      />
                    </div>
                  </div>
                )}

                {/* Pulsanti azioni */}
                <div className="flex space-x-2 pt-3">
                  <Button
                    onClick={resetForm}
                    variant="outline"
                    className="flex-1 border-slate-600 text-slate-200 hover:bg-slate-700"
                  >
                    Reset
                  </Button>
                  <Button
                    onClick={generateAndShowModal}
                    disabled={!canGenerateCard()}
                    className="flex-1 bg-lime-400 hover:bg-lime-500 text-slate-900 font-semibold"
                  >
                    <ImageIcon className="mr-2 h-4 w-4" />
                    Crea la mia Card
                  </Button>
                </div>
              </CardContent>
            </Card>

            {/* Anteprima Card - ora usa InstagramStoryTemplate */}
            <div className="flex flex-col lg:col-span-2">
              <Label className="text-slate-200 mb-3 text-lg font-semibold">Anteprima Card</Label>
              <div className="flex-1 bg-slate-900 rounded-lg border border-slate-700 p-4 min-h-[400px] overflow-hidden">
                {canGenerateCard() ? (
                  <div className="flex justify-start">
                    <div className="transform scale-[0.22] origin-top-left">
                      <InstagramStoryTemplate
                        athleteName={athleteName}
                        country={selectedCountry!}
                        profileImage={profileImage}
                        profile={athleteProfile}
                        isRunning={false}
                      />
                    </div>
                  </div>
                ) : (
                  <div className="flex items-center justify-center h-full">
                    <div className="text-center text-slate-400">
                      <ImageIcon className="w-16 h-16 mx-auto mb-4 opacity-50" />
                      <p>Compila tutti i campi per vedere</p>
                      <p>l'anteprima della tua card</p>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        </DialogContent>
      </Dialog>

      {/* Image Cropper Modal */}
      <ImageCropper
        isOpen={showImageCropper}
        onClose={() => setShowImageCropper(false)}
        imageSrc={tempImageSrc}
        onCropComplete={handleCropComplete}
      />

      {/* Social Share Modal */}
      {showSocialModal && canGenerateCard() && (
        <SocialShareModal
          isOpen={showSocialModal}
          onClose={() => setShowSocialModal(false)}
          athleteName={athleteName}
          country={selectedCountry!}
          profileImage={profileImage}
          profile={athleteProfile}
          isRunning={false}
        />
      )}
    </>
  );
};

export default AthleteCardGenerator;
