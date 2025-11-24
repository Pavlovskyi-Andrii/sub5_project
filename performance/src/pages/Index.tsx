
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Link } from 'react-router-dom';
import { BookOpen } from 'lucide-react';
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip';
import { useState } from 'react';
import { useTranslation } from 'react-i18next';
import PowerCurveModeler from '../components/PowerCurveModeler';
import BrandingFooter from '../components/BrandingFooter';
import GlossaryModal from '../components/GlossaryModal';
import LanguageSelector from '../components/LanguageSelector';

const Index = () => {
  const [glossaryOpen, setGlossaryOpen] = useState(false);
  const { t } = useTranslation();

  return (
    <div className="min-h-screen bg-slate-900">
      {/* Enhanced Navigation */}
      <div className="bg-slate-800 border-b border-slate-700 p-4">
        <div className="max-w-7xl mx-auto flex justify-between items-center">
          <div className="flex items-center gap-3">
            <div>
              <h1 className="text-2xl font-bold text-lime-400">{t('app_title')}</h1>
              <p className="text-slate-400 text-sm">{t('app_subtitle_cycling')}</p>
            </div>
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  onClick={() => setGlossaryOpen(true)}
                  className="text-lime-400 hover:text-lime-300 hover:bg-slate-700 flex items-center gap-2 px-3 py-2"
                >
                  <BookOpen className="h-5 w-5" />
                  <span className="text-sm font-medium">{t('nav_wiki')}</span>
                </Button>
              </TooltipTrigger>
              <TooltipContent>
                <p>{t('tooltip_wiki')}</p>
              </TooltipContent>
            </Tooltip>
          </div>
          <div className="flex items-center space-x-4">
            <LanguageSelector />
            <Button asChild variant="outline" className="border-lime-400 text-lime-400 hover:bg-lime-400 hover:text-slate-900">
              <Link to="/">{t('nav_cycling')}</Link>
            </Button>
            <Button asChild className="bg-lime-400 text-slate-900 hover:bg-lime-500">
              <Link to="/running">{t('nav_running')}</Link>
            </Button>
          </div>
        </div>
      </div>
      
      <PowerCurveModeler />
      <BrandingFooter />
      <GlossaryModal open={glossaryOpen} onOpenChange={setGlossaryOpen} />
    </div>
  );
};

export default Index;
