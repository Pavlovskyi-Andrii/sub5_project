import React from 'react';
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogClose,
} from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { ScrollArea } from "@/components/ui/scroll-area";
import { X } from "lucide-react";
import { useTranslation } from 'react-i18next';

interface GlossaryModalProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

const GlossaryModal: React.FC<GlossaryModalProps> = ({ open, onOpenChange }) => {
  const { t } = useTranslation();
  
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-4xl h-[90vh] bg-slate-900 border-slate-700 text-slate-100">
        <DialogHeader className="border-b border-slate-700 pb-4">
          <DialogTitle className="text-2xl font-bold text-lime-400">
            {t('wiki.title')}
          </DialogTitle>
          <p className="text-slate-300 text-base mt-3 leading-relaxed mb-4">
            {t('wiki.intro')}
          </p>
        </DialogHeader>
        
        <ScrollArea className="h-full pr-4 [&>[data-radix-scroll-area-viewport]]:scrollbar-thin [&>[data-radix-scroll-area-viewport]]:scrollbar-track-slate-800 [&>[data-radix-scroll-area-viewport]]:scrollbar-thumb-slate-600">
          <div className="space-y-12 py-6 text-left">
            
            {/* Sezione 1: Modelli Fisiologici */}
            <section>
              <h2 className="text-2xl font-bold text-white mb-6 mt-8">{t('wiki.s1.title')}</h2>
              <p className="text-slate-200 text-base mb-6 leading-7">
                {t('wiki.s1.intro')}
              </p>

              {/* Modello CP/W' */}
              <div className="mb-10">
                <h3 className="text-xl font-semibold text-lime-400 mb-4 mt-6">{t('wiki.s1.cpw.title')}</h3>
                <div className="bg-slate-800 p-6 rounded-lg border border-slate-700 space-y-4">
                  <div 
                    className="text-base leading-7 mb-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s1.cpw.what') }} 
                  />
                  <div 
                    className="text-base leading-7 mb-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s1.cpw.measures') }} 
                  />
                  
                  <div className="pl-6 border-l-2 border-lime-400 space-y-3">
                    <div 
                      className="text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s1.cpw.cp_definition') }} 
                    />
                    <div 
                      className="text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s1.cpw.w_prime_definition') }} 
                    />
                  </div>

                  <div className="grid md:grid-cols-2 gap-6 mt-6">
                    <div>
                      <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s1.cpw.strengths_title')}</h4>
                      <ul className="text-slate-200 text-base space-y-2 leading-7">
                        {Array.isArray(t('wiki.s1.cpw.strengths', { returnObjects: true })) && 
                          (t('wiki.s1.cpw.strengths', { returnObjects: true }) as string[]).map((strength, index) => (
                            <li key={index} dangerouslySetInnerHTML={{ __html: `• ${strength}` }} />
                          ))
                        }
                      </ul>
                    </div>
                    <div>
                      <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s1.cpw.weaknesses_title')}</h4>
                      <ul className="text-slate-200 text-base space-y-2 leading-7">
                        {Array.isArray(t('wiki.s1.cpw.weaknesses', { returnObjects: true })) && 
                          (t('wiki.s1.cpw.weaknesses', { returnObjects: true }) as string[]).map((weakness, index) => (
                            <li key={index} dangerouslySetInnerHTML={{ __html: `• ${weakness}` }} />
                          ))
                        }
                      </ul>
                    </div>
                  </div>
                  
                  <div 
                    className="text-base leading-7 mb-4 mt-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s1.cpw.integration') }} 
                  />
                </div>
              </div>

              {/* Modello APR */}
              <div className="mb-10">
                <h3 className="text-xl font-semibold text-lime-400 mb-4 mt-6">{t('wiki.s1.apr.title')}</h3>
                <div className="bg-slate-800 p-6 rounded-lg border border-slate-700 space-y-4">
                  <div 
                    className="text-base leading-7 mb-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s1.apr.what') }} 
                  />
                  <div 
                    className="text-base leading-7 mb-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s1.apr.measures') }} 
                  />
                  
                  <div className="pl-6 border-l-2 border-lime-400 space-y-3">
                    <div 
                      className="text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s1.apr.pmax_definition') }} 
                    />
                    <div 
                      className="text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s1.apr.p3min_definition') }} 
                    />
                    <div 
                      className="text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s1.apr.apr_definition') }} 
                    />
                    <div 
                      className="text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s1.apr.k_definition') }} 
                    />
                  </div>

                  <div className="grid md:grid-cols-2 gap-6 mt-6">
                    <div>
                      <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s1.apr.strengths_title')}</h4>
                      <ul className="text-slate-200 text-base space-y-2 leading-7">
                        {Array.isArray(t('wiki.s1.apr.strengths', { returnObjects: true })) && 
                          (t('wiki.s1.apr.strengths', { returnObjects: true }) as string[]).map((strength, index) => (
                            <li key={index} dangerouslySetInnerHTML={{ __html: `• ${strength}` }} />
                          ))
                        }
                      </ul>
                    </div>
                    <div>
                      <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s1.apr.weaknesses_title')}</h4>
                      <ul className="text-slate-200 text-base space-y-2 leading-7">
                        {Array.isArray(t('wiki.s1.apr.weaknesses', { returnObjects: true })) && 
                          (t('wiki.s1.apr.weaknesses', { returnObjects: true }) as string[]).map((weakness, index) => (
                            <li key={index} dangerouslySetInnerHTML={{ __html: `• ${weakness}` }} />
                          ))
                        }
                      </ul>
                    </div>
                  </div>
                  
                  <div 
                    className="text-base leading-7 mb-4 mt-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s1.apr.integration') }} 
                  />
                </div>
              </div>

              {/* Modello Power Law */}
              <div className="mb-10">
                <h3 className="text-xl font-semibold text-lime-400 mb-4 mt-6">{t('wiki.s1.power_law.title')}</h3>
                <div className="bg-slate-800 p-6 rounded-lg border border-slate-700 space-y-4">
                  <div 
                    className="text-base leading-7 mb-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s1.power_law.what') }} 
                  />
                  <div 
                    className="text-base leading-7 mb-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s1.power_law.measures') }} 
                  />
                  
                  <div className="pl-6 border-l-2 border-lime-400 space-y-3">
                    <div 
                      className="text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s1.power_law.s_definition') }} 
                    />
                    <div 
                      className="text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s1.power_law.e_definition') }} 
                    />
                  </div>

                  <div className="grid md:grid-cols-2 gap-6 mt-6">
                    <div>
                      <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s1.power_law.strengths_title')}</h4>
                      <ul className="text-slate-200 text-base space-y-2 leading-7">
                        {Array.isArray(t('wiki.s1.power_law.strengths', { returnObjects: true })) && 
                          (t('wiki.s1.power_law.strengths', { returnObjects: true }) as string[]).map((strength, index) => (
                            <li key={index} dangerouslySetInnerHTML={{ __html: `• ${strength}` }} />
                          ))
                        }
                      </ul>
                    </div>
                    <div>
                      <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s1.power_law.weaknesses_title')}</h4>
                      <ul className="text-slate-200 text-base space-y-2 leading-7">
                        {Array.isArray(t('wiki.s1.power_law.weaknesses', { returnObjects: true })) && 
                          (t('wiki.s1.power_law.weaknesses', { returnObjects: true }) as string[]).map((weakness, index) => (
                            <li key={index} dangerouslySetInnerHTML={{ __html: `• ${weakness}` }} />
                          ))
                        }
                      </ul>
                    </div>
                  </div>

                  <div 
                    className="text-base leading-7 mb-4 mt-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s1.power_law.integration') }} 
                  />
                </div>
              </div>
            </section>

            <hr className="border-slate-700" />

            {/* Sezione 2: Parametri dei Modelli */}
            <section>
              <h2 className="text-2xl font-bold text-white mb-6 mt-8">{t('wiki.s2.title')}</h2>
              
              {/* Parametri CP/W' */}
              <div className="mb-8">
                <h3 className="text-xl font-semibold text-lime-400 mb-4 mt-6">{t('wiki.s2.cp_w.title')}</h3>
                <div className="space-y-6">
                  <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                    <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s2.cp_w.cp.title')}</h4>
                    <p 
                      className="text-slate-200 text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.cp_w.cp.explanation') }} 
                    />
                    <p 
                      className="text-slate-200 text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.cp_w.cp.units') }} 
                    />
                    <p 
                      className="text-slate-200 text-base leading-7" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.cp_w.cp.calculation') }} 
                    />
                  </div>
                  
                  <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                    <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s2.cp_w.w_prime.title')}</h4>
                    <p 
                      className="text-slate-200 text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.cp_w.w_prime.explanation') }} 
                    />
                    <p 
                      className="text-slate-200 text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.cp_w.w_prime.units') }} 
                    />
                    <p 
                      className="text-slate-200 text-base leading-7" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.cp_w.w_prime.calculation') }} 
                    />
                  </div>
                </div>
              </div>

              {/* Parametri APR */}
              <div className="mb-8">
                <h3 className="text-xl font-semibold text-lime-400 mb-4 mt-6">{t('wiki.s2.apr.title')}</h3>
                <div className="space-y-6">
                  <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                    <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s2.apr.pmax.title')}</h4>
                    <p 
                      className="text-slate-200 text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.apr.pmax.explanation') }} 
                    />
                    <p 
                      className="text-slate-200 text-base leading-7" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.apr.pmax.units') }} 
                    />
                  </div>
                  
                  <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                    <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s2.apr.p3min.title')}</h4>
                    <p 
                      className="text-slate-200 text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.apr.p3min.explanation') }} 
                    />
                    <p 
                      className="text-slate-200 text-base leading-7" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.apr.p3min.units') }} 
                    />
                  </div>
                  
                  <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                    <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s2.apr.apr_param.title')}</h4>
                    <p 
                      className="text-slate-200 text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.apr.apr_param.explanation') }} 
                    />
                    <p 
                      className="text-slate-200 text-base leading-7" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.apr.apr_param.units') }} 
                    />
                  </div>

                  <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                    <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s2.apr.k.title')}</h4>
                    <p 
                      className="text-slate-200 text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.apr.k.explanation') }} 
                    />
                    <p 
                      className="text-slate-200 text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.apr.k.units') }} 
                    />
                    <p 
                      className="text-slate-200 text-base leading-7" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.apr.k.calculation') }} 
                    />
                  </div>
                </div>
              </div>

              {/* Parametri Power Law */}
              <div className="mb-8">
                <h3 className="text-xl font-semibold text-lime-400 mb-4 mt-6">{t('wiki.s2.power_law.title')}</h3>
                <div className="space-y-6">
                  <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                    <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s2.power_law.s.title')}</h4>
                    <p 
                      className="text-slate-200 text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.power_law.s.explanation') }} 
                    />
                    <p 
                      className="text-slate-200 text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.power_law.s.units') }} 
                    />
                    <p 
                      className="text-slate-200 text-base leading-7" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.power_law.s.calculation') }} 
                    />
                  </div>
                  
                  <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                    <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s2.power_law.e.title')}</h4>
                    <p 
                      className="text-slate-200 text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.power_law.e.explanation') }} 
                    />
                    <p 
                      className="text-slate-200 text-base leading-7 mb-4" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.power_law.e.units') }} 
                    />
                    <p 
                      className="text-slate-200 text-base leading-7" 
                      dangerouslySetInnerHTML={{ __html: t('wiki.s2.power_law.e.calculation') }} 
                    />
                  </div>
                </div>
              </div>
            </section>

            <hr className="border-slate-700" />

            {/* Sezione 3: Formule */}
            <section>
              <h2 className="text-2xl font-bold text-white mb-6 mt-8">{t('wiki.s3.title')}</h2>
              <p className="text-slate-200 text-base leading-7 mb-8">{t('wiki.s3.intro')}</p>
              
              <div className="space-y-8">
                <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                  <h3 className="text-xl font-semibold text-lime-400 mb-4">{t('wiki.s3.cp_w_formula.title')}</h3>
                  <div 
                    className="bg-slate-900 p-4 rounded font-mono text-lime-400 mb-4 text-base"
                    dangerouslySetInnerHTML={{ __html: t('wiki.s3.cp_w_formula.formula') }}
                  />
                </div>

                <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                  <h3 className="text-xl font-semibold text-lime-400 mb-4">{t('wiki.s3.apr_formula.title')}</h3>
                  <div 
                    className="bg-slate-900 p-4 rounded font-mono text-lime-400 mb-4 text-base"
                    dangerouslySetInnerHTML={{ __html: t('wiki.s3.apr_formula.formula') }}
                  />
                </div>

                <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                  <h3 className="text-xl font-semibold text-lime-400 mb-4">{t('wiki.s3.power_law_formula.title')}</h3>
                  <div 
                    className="bg-slate-900 p-4 rounded font-mono text-lime-400 mb-4 text-base"
                    dangerouslySetInnerHTML={{ __html: t('wiki.s3.power_law_formula.formula') }}
                  />
                </div>
              </div>
            </section>

            <hr className="border-slate-700" />

            {/* Sezione 4: Punteggi */}
            <section>
              <h2 className="text-2xl font-bold text-white mb-6 mt-8">{t('wiki.s4.title')}</h2>
              <p className="text-slate-200 text-base leading-7 mb-8">{t('wiki.s4.intro')}</p>
              
              <div className="mb-8">
                <h3 className="text-xl font-semibold text-lime-400 mb-4 mt-6">{t('wiki.s4.calculations.title')}</h3>
                <div className="bg-slate-800 p-6 rounded-lg border border-slate-700 space-y-4">
                  <p 
                    className="text-base leading-7 mb-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s4.calculations.what_are') }} 
                  />
                  <p 
                    className="text-base leading-7 mb-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s4.calculations.differentiated') }} 
                  />
                  
                  <div className="grid md:grid-cols-2 gap-6">
                    <div className="bg-slate-900 p-4 rounded border border-slate-600">
                      <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s4.calculations.cycling_formula.title')}</h4>
                      <div className="font-mono text-lime-400 mb-2">{t('wiki.s4.calculations.cycling_formula.formula')}</div>
                      <p className="text-slate-200 text-sm">{t('wiki.s4.calculations.cycling_formula.explanation')}</p>
                    </div>
                    <div className="bg-slate-900 p-4 rounded border border-slate-600">
                      <h4 className="font-bold text-lime-400 mb-3">{t('wiki.s4.calculations.running_formula.title')}</h4>
                      <div className="font-mono text-lime-400 mb-2">{t('wiki.s4.calculations.running_formula.formula')}</div>
                      <p className="text-slate-200 text-sm">{t('wiki.s4.calculations.running_formula.explanation')}</p>
                    </div>
                  </div>
                  
                  <p 
                    className="text-base leading-7 mb-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s4.calculations.rounding') }} 
                  />
                  <p 
                    className="text-base leading-7 mb-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s4.calculations.epic_effect') }} 
                  />
                </div>
              </div>

              <div className="grid md:grid-cols-2 gap-8 mb-8">
                <div>
                  <h3 className="text-xl font-semibold text-lime-400 mb-4 mt-6">{t('wiki.s4.cycling_spider.title')}</h3>
                  <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                    <p className="text-slate-200 text-base leading-7 mb-4">{t('wiki.s4.cycling_spider.intro')}</p>
                    <ul className="text-slate-200 text-base space-y-2 leading-7">
                      {
                        Array.isArray(t('wiki.s4.cycling_spider.stats', { returnObjects: true })) ? 
                        (t('wiki.s4.cycling_spider.stats', { returnObjects: true }) as string[]).map((stat, index) => (
                          <li key={index} dangerouslySetInnerHTML={{ __html: stat }} />
                        )) : []
                      }
                    </ul>
                  </div>
                </div>

                <div>
                  <h3 className="text-xl font-semibold text-lime-400 mb-4 mt-6">{t('wiki.s4.running_spider.title')}</h3>
                  <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                    <p className="text-slate-200 text-base leading-7 mb-4">{t('wiki.s4.running_spider.intro')}</p>
                    <ul className="text-slate-200 text-base space-y-2 leading-7">
                      {
                        Array.isArray(t('wiki.s4.running_spider.stats', { returnObjects: true })) ? 
                        (t('wiki.s4.running_spider.stats', { returnObjects: true }) as string[]).map((stat, index) => (
                          <li key={index} dangerouslySetInnerHTML={{ __html: stat }} />
                        )) : []
                      }
                    </ul>
                  </div>
                </div>
              </div>

              <div className="mb-8">
                <h3 className="text-xl font-semibold text-lime-400 mb-4 mt-6">{t('wiki.s4.overall_rating.title')}</h3>
                <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                  <p 
                    className="text-base leading-7 mb-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s4.overall_rating.what_is') }} 
                  />
                  <p 
                    className="text-base leading-7 mb-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s4.overall_rating.rarity') }} 
                  />
                </div>
              </div>

              <div className="mb-8">
                <h3 className="text-xl font-semibold text-lime-400 mb-4 mt-6">{t('wiki.s4.athlete_profile.title')}</h3>
                <div className="bg-slate-800 p-6 rounded-lg border border-slate-700">
                  <p 
                    className="text-base leading-7 mb-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s4.athlete_profile.what_is') }} 
                  />
                  <p 
                    className="text-base leading-7 mb-4" 
                    dangerouslySetInnerHTML={{ __html: t('wiki.s4.athlete_profile.how_works') }} 
                  />
                </div>
              </div>
            </section>

            <hr className="border-slate-700" />

            {/* Bibliografia */}
            <section>
              <h2 className="text-2xl font-bold text-white mb-6 mt-8">{t('wiki.s5.title')}</h2>
              <div className="bg-slate-800 p-6 rounded-lg border border-slate-700 space-y-4">
                <p className="text-slate-200 text-base leading-7 mb-4">{t('wiki.s5.intro')}</p>
                <ul className="text-slate-200 text-sm space-y-3 leading-6">
                  {
                    Array.isArray(t('wiki.s5.references', { returnObjects: true })) ? 
                    (t('wiki.s5.references', { returnObjects: true }) as string[]).map((reference, index) => (
                      <li key={index} className="text-slate-200 text-sm leading-6">{reference}</li>
                    )) : []
                  }
                </ul>
              </div>
            </section>
          </div>
        </ScrollArea>
        
        <DialogClose className="absolute right-4 top-4 rounded-sm opacity-70 ring-offset-background transition-opacity hover:opacity-100 focus:outline-none focus:ring-2 focus:ring-ring focus:ring-offset-2 disabled:pointer-events-none data-[state=open]:bg-accent data-[state=open]:text-muted-foreground">
          <X className="h-5 w-5 text-slate-400 hover:text-lime-400" />
          <span className="sr-only">Close</span>
        </DialogClose>
      </DialogContent>
    </Dialog>
  );
};

export default GlossaryModal;