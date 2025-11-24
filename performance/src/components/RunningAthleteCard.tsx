import React from 'react';
import { RunningAthleteProfile, RUNNING_STAT_METADATA, RarityTier } from '@/utils/runningAthleteCalculations';
import { Country } from '@/utils/countries';
import miccoachLogo from '@/assets/miccoach-logo.png';
import { useTranslation } from 'react-i18next';

interface RunningAthleteCardProps {
  athleteName: string;
  country: Country;
  profileImage: string;
  profile: RunningAthleteProfile;
  format?: 'square' | 'story';
}

// Stili per i tier di rarità (identici alla versione ciclismo)
const getTierStyles = (tier: RarityTier) => {
  switch (tier) {
    case 'ALIEN':
      return {
        ratingBg: 'bg-gradient-to-br from-purple-500 via-pink-500 to-cyan-500',
        ratingBorder: 'border-purple-500',
        photoBorder: 'border-purple-500 shadow-purple-500/50',
        animation: 'animate-pulse'
      };
    case 'HERO':
      return {
        ratingBg: 'bg-gradient-to-br from-yellow-400 to-orange-500',
        ratingBorder: 'border-yellow-400',
        photoBorder: 'border-yellow-400 shadow-yellow-400/50',
        animation: ''
      };
    case 'PRO':
      return {
        ratingBg: 'bg-gradient-to-br from-purple-400 to-violet-600',
        ratingBorder: 'border-purple-400',
        photoBorder: 'border-purple-400 shadow-purple-400/50',
        animation: ''
      };
    case 'ELITE':
      return {
        ratingBg: 'bg-gradient-to-br from-blue-400 to-blue-600',
        ratingBorder: 'border-blue-400',
        photoBorder: 'border-blue-400 shadow-blue-400/50',
        animation: ''
      };
    default: // STANDARD
      return {
        ratingBg: 'bg-gradient-to-br from-lime-400/20 to-green-600/20',
        ratingBorder: 'border-lime-400',
        photoBorder: 'border-lime-400 shadow-lime-400/50',
        animation: ''
      };
  }
};

const RunningAthleteCard: React.FC<RunningAthleteCardProps> = ({
  athleteName,
  country,
  profileImage,
  profile,
  format = 'square'
}) => {
  const { t } = useTranslation();
  const { stats, overallRating, profileType, rarityTier } = profile;
  const tierStyles = getTierStyles(rarityTier);

  // Get translated profile name
  const getProfileName = () => {
    const profileKey = profileType.toLowerCase().replace(/[\s-]/g, '_');
    return t(`athlete_card.profiles.running.${profileKey}`, profileType);
  };

  // Dimensioni in base al formato
  const isStory = format === 'story';
  const cardDimensions = isStory 
    ? { width: '1080px', height: '1920px' } 
    : { width: '1080px', height: '1080px' };

  if (isStory) {
    // Layout verticale per Instagram Stories
    return (
      <div
        className="relative text-white overflow-hidden"
        style={{
          ...cardDimensions,
          background: 'radial-gradient(circle at center, #2a2a2e 0%, #1a1a1c 100%)',
          position: 'relative'
        }}
      >
        {/* Pattern geometrico di sfondo */}
        <div 
          className="absolute inset-0 opacity-5"
          style={{
            backgroundImage: `url("data:image/svg+xml,%3Csvg width='60' height='60' viewBox='0 0 60 60' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' fill-rule='evenodd'%3E%3Cg fill='%23ffffff' fill-opacity='0.1'%3E%3Cpath d='m36 34v-4h-2v4h-4v2h4v4h2v-4h4v-2h-4zm0-30V0h-2v4h-4v2h4v4h2V6h4V4h-4zM6 34v-4H4v4H0v2h4v4h2v-4h4v-2H6zM6 4V0H4v4H0v2h4v4h2V6h4V4H6z'/%3E%3C/g%3E%3C/g%3E%3C/svg%3E")`,
          }}
        />

        {/* Logo MicCoach Lab */}
        <div className="absolute top-12 left-8 z-10">
          <img
            src={miccoachLogo}
            alt="MicCoach Lab"
            className="h-20 w-auto opacity-90"
          />
        </div>

        {/* Foto atleta grande in alto */}
        <div className="absolute top-32 left-1/2 transform -translate-x-1/2">
          <div className={`w-80 h-80 rounded-full overflow-hidden border-6 ${tierStyles.photoBorder} shadow-2xl ${tierStyles.animation} relative`}>
            <div className="absolute inset-0 bg-gradient-to-br from-lime-400/20 to-transparent rounded-full"></div>
            <img
              src={profileImage}
              alt={athleteName}
              className="w-full h-full object-cover relative z-10"
            />
          </div>
        </div>

        {/* Nome e nazionalità */}
        <div className="absolute top-[480px] left-8 right-8 text-center">
          <div className="flex items-center justify-center space-x-6 mb-4">
            <span className="text-8xl">{country.flag}</span>
          </div>
          <h1 className="text-6xl font-black text-lime-400 leading-tight mb-2" style={{ fontFamily: 'Inter, sans-serif', fontWeight: 900 }}>
            {athleteName}
          </h1>
          <p className="text-3xl text-slate-300 font-semibold">
            {country.name}
          </p>
        </div>

        {/* Rating complessivo */}
        <div className="absolute top-[730px] left-1/2 transform -translate-x-1/2">
          <div className="relative">
            {/* SVG Background per garantire la forma ottagonale */}
            <svg 
              width="280" 
              height="200" 
              viewBox="0 0 280 200" 
              className="absolute inset-0"
            >
              <defs>
                <linearGradient id={`running-gradient-${rarityTier}-story`} x1="0%" y1="0%" x2="100%" y2="100%">
                  {rarityTier === 'ALIEN' && (
                    <>
                      <stop offset="0%" stopColor="#a855f7" />
                      <stop offset="50%" stopColor="#ec4899" />
                      <stop offset="100%" stopColor="#06b6d4" />
                    </>
                  )}
                  {rarityTier === 'HERO' && (
                    <>
                      <stop offset="0%" stopColor="#facc15" />
                      <stop offset="100%" stopColor="#f97316" />
                    </>
                  )}
                  {rarityTier === 'PRO' && (
                    <>
                      <stop offset="0%" stopColor="#a78bfa" />
                      <stop offset="100%" stopColor="#7c3aed" />
                    </>
                  )}
                  {rarityTier === 'ELITE' && (
                    <>
                      <stop offset="0%" stopColor="#60a5fa" />
                      <stop offset="100%" stopColor="#2563eb" />
                    </>
                  )}
                  {rarityTier === 'STANDARD' && (
                    <>
                      <stop offset="0%" stopColor="rgba(163, 230, 53, 0.2)" />
                      <stop offset="100%" stopColor="rgba(22, 163, 74, 0.2)" />
                    </>
                  )}
                </linearGradient>
              </defs>
              <polygon 
                points="56,0 224,0 280,60 280,140 224,200 56,200 0,140 0,60"
                fill={`url(#running-gradient-${rarityTier}-story)`}
                stroke={tierStyles.ratingBorder.includes('purple') ? '#a855f7' : 
                       tierStyles.ratingBorder.includes('yellow') ? '#facc15' :
                       tierStyles.ratingBorder.includes('blue') ? '#60a5fa' :
                       '#a3e635'}
                strokeWidth="4"
                className={tierStyles.animation}
              />
            </svg>
            
            {/* Contenuto del rating */}
            <div className="relative z-10 flex items-center justify-center h-[200px] w-[280px]">
              <div className="text-center">
                <div className="text-7xl font-black text-white mb-4" style={{ fontFamily: 'Inter, sans-serif', fontWeight: 900 }}>
                  {overallRating}
                </div>
                <div className="text-lg font-bold text-white uppercase tracking-widest">
                  {getProfileName()}
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Statistiche verticali */}
        <div className="absolute top-[1000px] left-8 right-8 bottom-32 space-y-4">
          {Object.entries(stats).map(([key, value]) => {
            // Map keys for professional names
            const displayKey = key === 'Mezza' ? t('athlete_card.distances.half_marathon') : 
                             key === 'Mara' ? t('athlete_card.distances.marathon') : key;
            
            return (
              <div 
                key={key} 
                className="bg-gradient-to-br from-slate-800/80 to-slate-900/80 backdrop-blur-sm border border-slate-700/50 p-6 relative overflow-hidden"
                style={{
                  clipPath: 'polygon(0 0, calc(100% - 15px) 0, 100% 15px, 100% 100%, 0 100%)'
                }}
              >
                {/* Effetto luminoso interno */}
                <div className="absolute top-0 right-0 w-16 h-16 bg-lime-400/5 blur-xl"></div>
                
                <div className="flex justify-between items-center relative z-10">
                  <div className="flex items-center space-x-4">
                    <div>
                      <div className="text-2xl font-black text-lime-400" style={{ fontFamily: 'Inter, sans-serif', fontWeight: 900 }}>
                        {displayKey}
                      </div>
                    </div>
                  </div>
                  <div className={`text-4xl font-black ${value >= 100 ? 'text-yellow-400' : 'text-white'}`} style={{ fontFamily: 'Inter, sans-serif', fontWeight: 900 }}>
                    {value}
                  </div>
                </div>
                
                {/* Barra di progresso */}
                <div className="mt-4">
                  <div className="h-3 bg-slate-700 rounded-full overflow-hidden">
                    <div 
                      className="h-full rounded-full transition-all duration-1000 ease-out"
                      style={{
                        width: `${Math.min(value, 100)}%`,
                        background: value >= 100 
                          ? 'linear-gradient(to right, #F97316, #DC2626)' 
                          : 'linear-gradient(to right, #FACC15, #4ADE80)'
                      }}
                    />
                  </div>
                </div>
              </div>
            );
          })}
        </div>

        {/* Footer */}
        <div className="absolute bottom-8 left-8 right-8 text-center">
          <p className="text-slate-400 text-xl font-medium">
            {t('athlete_card.generated_on')} <span className="text-lime-400 font-bold">miccoach.it</span>
          </p>
        </div>

        {/* Effetti decorativi */}
        <div className="absolute -top-20 -right-20 w-40 h-40 bg-lime-400/10 rounded-full blur-3xl"></div>
        <div className="absolute -bottom-20 -left-20 w-60 h-60 bg-lime-400/5 rounded-full blur-3xl"></div>
      </div>
    );
  }

  // Layout quadrato per i post
  return (
    <div
      className="relative text-white overflow-hidden"
      style={{
        ...cardDimensions,
        background: 'radial-gradient(circle at center, #2a2a2e 0%, #1a1a1c 100%)',
        position: 'relative'
      }}
    >
      {/* Pattern geometrico di sfondo */}
      <div 
        className="absolute inset-0 opacity-5"
        style={{
          backgroundImage: `url("data:image/svg+xml,%3Csvg width='60' height='60' viewBox='0 0 60 60' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' fill-rule='evenodd'%3E%3Cg fill='%23ffffff' fill-opacity='0.1'%3E%3Cpath d='m36 34v-4h-2v4h-4v2h4v4h2v-4h4v-2h-4zm0-30V0h-2v4h-4v2h4v4h2V6h4V4h-4zM6 34v-4H4v4H0v2h4v4h2v-4h4v-2H6zM6 4V0H4v4H0v2h4v4h2V6h4V4H6z'/%3E%3C/g%3E%3C/g%3E%3C/svg%3E")`,
        }}
      />

      {/* Logo MicCoach Lab */}
      <div className="absolute top-8 left-8 z-10">
        <img
          src={miccoachLogo}
          alt="MicCoach Lab"
          className="h-16 w-auto opacity-90"
        />
      </div>

      {/* Rating complessivo - Stile scudo/esagono */}
      <div className="absolute top-8 right-8 z-10">
        <div className="relative">
          {/* SVG Background per garantire la forma ottagonale */}
          <svg 
            width="200" 
            height="140" 
            viewBox="0 0 200 140" 
            className="absolute inset-0"
          >
            <defs>
              <linearGradient id={`running-gradient-${rarityTier}-square`} x1="0%" y1="0%" x2="100%" y2="100%">
                {rarityTier === 'ALIEN' && (
                  <>
                    <stop offset="0%" stopColor="#a855f7" />
                    <stop offset="50%" stopColor="#ec4899" />
                    <stop offset="100%" stopColor="#06b6d4" />
                  </>
                )}
                {rarityTier === 'HERO' && (
                  <>
                    <stop offset="0%" stopColor="#facc15" />
                    <stop offset="100%" stopColor="#f97316" />
                  </>
                )}
                {rarityTier === 'PRO' && (
                  <>
                    <stop offset="0%" stopColor="#a78bfa" />
                    <stop offset="100%" stopColor="#7c3aed" />
                  </>
                )}
                {rarityTier === 'ELITE' && (
                  <>
                    <stop offset="0%" stopColor="#60a5fa" />
                    <stop offset="100%" stopColor="#2563eb" />
                  </>
                )}
                {rarityTier === 'STANDARD' && (
                  <>
                    <stop offset="0%" stopColor="rgba(163, 230, 53, 0.2)" />
                    <stop offset="100%" stopColor="rgba(22, 163, 74, 0.2)" />
                  </>
                )}
              </linearGradient>
            </defs>
            <polygon 
              points="40,0 160,0 200,42 200,98 160,140 40,140 0,98 0,42"
              fill={`url(#running-gradient-${rarityTier}-square)`}
              stroke={tierStyles.ratingBorder.includes('purple') ? '#a855f7' : 
                     tierStyles.ratingBorder.includes('yellow') ? '#facc15' :
                     tierStyles.ratingBorder.includes('blue') ? '#60a5fa' :
                     '#a3e635'}
              strokeWidth="2"
              className={tierStyles.animation}
            />
          </svg>
          
          {/* Contenuto del rating */}
          <div className="relative z-10 flex items-center justify-center h-[140px] w-[200px]">
            <div className="text-center">
              <div className="text-5xl font-black text-white mb-3" style={{ fontFamily: 'Inter, sans-serif', fontWeight: 900 }}>
                {overallRating}
              </div>
              <div className="text-sm font-bold text-white uppercase tracking-widest">
                {getProfileName()}
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Header con foto e info atleta */}
      <div className="absolute top-32 left-8 right-8 flex items-center space-x-8">
        {/* Foto atleta con bordo luminoso */}
        <div className="relative">
          <div className={`w-48 h-48 rounded-full overflow-hidden border-4 ${tierStyles.photoBorder} shadow-2xl ${tierStyles.animation} relative`}>
            <div className="absolute inset-0 bg-gradient-to-br from-lime-400/20 to-transparent rounded-full"></div>
            <img
              src={profileImage}
              alt={athleteName}
              className="w-full h-full object-cover relative z-10"
            />
          </div>
        </div>

        {/* Info atleta */}
        <div className="flex-1 space-y-4">
          <div className="flex items-center space-x-4">
            <span className="text-6xl">{country.flag}</span>
            <div>
              <h1 className="text-5xl font-black text-lime-400 leading-tight" style={{ fontFamily: 'Inter, sans-serif', fontWeight: 900 }}>
                {athleteName}
              </h1>
              <p className="text-2xl text-slate-300 mt-2 font-semibold">
                {country.name}
              </p>
            </div>
          </div>
        </div>
      </div>

      {/* Statistiche - 5 box in griglia 2x2 + 1 */}
      <div className="absolute top-96 left-8 right-8 bottom-24">
        <h2 className="text-3xl font-black text-lime-400 mb-8 text-center uppercase tracking-wider">
          {t('athlete_card.performance_analysis')}
        </h2>
        
        {/* Prima riga - 2 statistiche */}
        <div className="grid grid-cols-2 gap-8 mb-8">
          {Object.entries(stats).slice(0, 2).map(([key, value]) => {
            // Map keys for professional names
            const displayKey = key === 'Mezza' ? t('athlete_card.distances.half_marathon') : 
                             key === 'Mara' ? t('athlete_card.distances.marathon') : key;
            
            return (
              <div 
                key={key} 
                className="bg-gradient-to-br from-slate-800/80 to-slate-900/80 backdrop-blur-sm border border-slate-700/50 p-8 relative overflow-hidden"
                style={{
                  clipPath: 'polygon(0 0, calc(100% - 20px) 0, 100% 20px, 100% 100%, 0 100%)'
                }}
              >
                {/* Effetto luminoso interno */}
                <div className="absolute top-0 right-0 w-20 h-20 bg-lime-400/5 blur-xl"></div>
                
                <div className="flex justify-between items-start mb-4 relative z-10">
                  <div>
                    <div className="text-3xl font-black text-lime-400" style={{ fontFamily: 'Inter, sans-serif', fontWeight: 900 }}>
                      {displayKey}
                    </div>
                  </div>
                  <div className={`text-5xl font-black ${value >= 100 ? 'text-yellow-400' : 'text-white'}`} style={{ fontFamily: 'Inter, sans-serif', fontWeight: 900 }}>
                    {value}
                  </div>
                </div>
                
                {/* Barra di progresso stilizzata */}
                <div className="relative">
                  <div className="h-2 bg-slate-700 rounded-full overflow-hidden">
                    <div 
                      className="h-full rounded-full transition-all duration-1000 ease-out"
                      style={{
                        width: `${Math.min(value, 100)}%`,
                        background: value >= 100 
                          ? 'linear-gradient(to right, #F97316, #DC2626)' 
                          : 'linear-gradient(to right, #FACC15, #4ADE80)'
                      }}
                    />
                  </div>
                </div>
              </div>
            );
          })}
        </div>

        {/* Seconda riga - 2 statistiche */}
        <div className="grid grid-cols-2 gap-8 mb-8">
          {Object.entries(stats).slice(2, 4).map(([key, value]) => {
            // Map keys for professional names
            const displayKey = key === 'Mezza' ? t('athlete_card.distances.half_marathon') : 
                             key === 'Mara' ? t('athlete_card.distances.marathon') : key;
            
            return (
              <div 
                key={key} 
                className="bg-gradient-to-br from-slate-800/80 to-slate-900/80 backdrop-blur-sm border border-slate-700/50 p-8 relative overflow-hidden"
                style={{
                  clipPath: 'polygon(0 0, calc(100% - 20px) 0, 100% 20px, 100% 100%, 0 100%)'
                }}
              >
                {/* Effetto luminoso interno */}
                <div className="absolute top-0 right-0 w-20 h-20 bg-lime-400/5 blur-xl"></div>
                
                <div className="flex justify-between items-start mb-4 relative z-10">
                  <div>
                    <div className="text-3xl font-black text-lime-400" style={{ fontFamily: 'Inter, sans-serif', fontWeight: 900 }}>
                      {displayKey}
                    </div>
                  </div>
                  <div className={`text-5xl font-black ${value >= 100 ? 'text-yellow-400' : 'text-white'}`} style={{ fontFamily: 'Inter, sans-serif', fontWeight: 900 }}>
                    {value}
                  </div>
                </div>
                
                {/* Barra di progresso stilizzata */}
                <div className="relative">
                  <div className="h-2 bg-slate-700 rounded-full overflow-hidden">
                    <div 
                      className="h-full rounded-full transition-all duration-1000 ease-out"
                      style={{
                        width: `${Math.min(value, 100)}%`,
                        background: value >= 100 
                          ? 'linear-gradient(to right, #F97316, #DC2626)' 
                          : 'linear-gradient(to right, #FACC15, #4ADE80)'
                      }}
                    />
                  </div>
                </div>
              </div>
            );
          })}
        </div>

        {/* Terza riga - 1 statistica centrata (Mara) */}
        <div className="flex justify-center">
          <div className="w-1/2">
            {Object.entries(stats).slice(4, 5).map(([key, value]) => {
              // Map keys for professional names
              const displayKey = key === 'Mezza' ? t('athlete_card.distances.half_marathon') : 
                               key === 'Mara' ? t('athlete_card.distances.marathon') : key;
              
              return (
                <div 
                  key={key} 
                  className="bg-gradient-to-br from-slate-800/80 to-slate-900/80 backdrop-blur-sm border border-slate-700/50 p-8 relative overflow-hidden"
                  style={{
                    clipPath: 'polygon(0 0, calc(100% - 20px) 0, 100% 20px, 100% 100%, 0 100%)'
                  }}
                >
                  {/* Effetto luminoso interno */}
                  <div className="absolute top-0 right-0 w-20 h-20 bg-lime-400/5 blur-xl"></div>
                  
                  <div className="flex justify-between items-start mb-4 relative z-10">
                    <div>
                      <div className="text-3xl font-black text-lime-400" style={{ fontFamily: 'Inter, sans-serif', fontWeight: 900 }}>
                        {displayKey}
                      </div>
                    </div>
                    <div className={`text-5xl font-black ${value >= 100 ? 'text-yellow-400' : 'text-white'}`} style={{ fontFamily: 'Inter, sans-serif', fontWeight: 900 }}>
                      {value}
                    </div>
                  </div>
                  
                  {/* Barra di progresso stilizzata */}
                  <div className="relative">
                    <div className="h-2 bg-slate-700 rounded-full overflow-hidden">
                      <div 
                        className="h-full rounded-full transition-all duration-1000 ease-out"
                        style={{
                          width: `${Math.min(value, 100)}%`,
                          background: value >= 100 
                            ? 'linear-gradient(to right, #F97316, #DC2626)' 
                            : 'linear-gradient(to right, #FACC15, #4ADE80)'
                        }}
                      />
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* Footer */}
      <div className="absolute bottom-6 left-8 right-8 text-center">
        <p className="text-slate-400 text-lg font-medium">
          {t('athlete_card.generated_on')} <span className="text-lime-400 font-bold">miccoach.it</span>
        </p>
      </div>

      {/* Effetti decorativi */}
      <div className="absolute -top-20 -right-20 w-40 h-40 bg-lime-400/10 rounded-full blur-3xl"></div>
      <div className="absolute -bottom-20 -left-20 w-60 h-60 bg-lime-400/5 rounded-full blur-3xl"></div>
      <div className="absolute top-1/2 left-1/2 transform -translate-x-1/2 -translate-y-1/2 w-96 h-96 bg-lime-400/3 rounded-full blur-3xl"></div>
    </div>
  );
};

export default RunningAthleteCard;