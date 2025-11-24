
import React from 'react';
import { useTranslation } from 'react-i18next';

const BrandingFooter = () => {
  const { t } = useTranslation();
  
  return (
    <footer className="bg-slate-900 border-t border-slate-700 py-6 mt-12">
      <div className="max-w-7xl mx-auto px-4">
        <div className="flex justify-between items-center flex-wrap gap-6">
          
          {/* Sezione Sinistra: Branding e Copyright */}
          <div className="text-left">
            <p className="text-lg font-bold text-white mb-1">{t('footer.brand')}</p>
            <p className="text-xs text-slate-400">{t('footer.copyright')}</p>
          </div>
          
          {/* Sezione Centrale: Tagline e Supporto */}
          <div className="text-center flex flex-col items-center gap-3">
            <p className="text-slate-300 italic">{t('footer.tagline')}</p>
            <a 
              href="https://buymeacoffee.com/michaelpesse" 
              target="_blank" 
              rel="noopener noreferrer"
              className="inline-flex items-center px-4 py-2 border border-lime-400 rounded text-lime-400 hover:bg-lime-400 hover:text-slate-900 transition-colors text-sm font-medium"
            >
              {t('footer.support_button')}
            </a>
            <p className="text-xs text-slate-400 max-w-xs">
              {t('footer.support_description')}
            </p>
          </div>
          
          {/* Sezione Destra: Contatti */}
          <div className="text-right">
            <a 
              href="mailto:info@miccoach.it" 
              className="text-slate-400 hover:text-lime-400 transition-colors text-sm block mb-1"
            >
              info@miccoach.it
            </a>
            <a 
              href="http://www.miccoach.it" 
              target="_blank" 
              rel="noopener noreferrer" 
              className="text-slate-400 hover:text-lime-400 transition-colors text-sm block"
            >
              www.miccoach.it
            </a>
          </div>
          
        </div>
      </div>
    </footer>
  );
};

export default BrandingFooter;
