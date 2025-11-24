import React from 'react';
import { Document, Page, Text, View, StyleSheet, pdf, Image } from '@react-pdf/renderer';
import { calculateAthleteProfile, formatProfileType, STAT_METADATA } from '@/utils/athleteCardCalculations';
import { PdfColorPalette, PDF_COLOR_PALETTES } from '@/utils/pdfColorPalettes';

interface UserData {
  nome: string;
  cognome: string;
  email: string;
  peso_kg: number;
}

interface PowerModels {
  valid: boolean;
  cp: number;
  wPrime: number;
  pMax: number;
  cp_power: number;
  apr_params: {
    po3min: number;
    apr: number;
    k: number;
  };
  pl_params: {
    S: number;
    E: number;
  };
  lt1_range: {
    min: number;
    max: number;
    estimate: number;
  };
  vo2max_power: number;
}

interface CoachData {
  isCoach: boolean;
  coachName?: string;
  coachCompany?: string;
  coachEmail?: string;
  coachLogo?: string;
}

interface PdfReportGeneratorProps {
  userData: UserData;
  models: PowerModels;
  chartImage?: string | null;
  profileImage?: string | null;
  coachData?: CoachData;
  colorPalette?: string;
}

const createStyles = (palette: PdfColorPalette) => StyleSheet.create({
  page: {
    flexDirection: 'column',
    backgroundColor: '#ffffff',
    padding: 30,
    fontFamily: 'Helvetica',
    fontSize: 10,
  },
  header: {
    alignItems: 'center',
    borderBottomWidth: 3,
    borderBottomColor: palette.border,
    paddingBottom: 15,
    marginBottom: 20,
  },
  title: {
    fontSize: 18,
    fontWeight: 'bold',
    color: palette.primary,
    marginBottom: 8,
    textAlign: 'center',
  },
  subtitle: {
    fontSize: 12,
    color: '#666666',
    textAlign: 'center',
  },
  personalInfo: {
    backgroundColor: palette.background,
    padding: 15,
    borderRadius: 8,
    borderLeftWidth: 4,
    borderLeftColor: palette.border,
    marginBottom: 20,
  },
  infoTitle: {
    fontSize: 14,
    fontWeight: 'bold',
    color: palette.primary,
    marginBottom: 8,
  },
  infoText: {
    fontSize: 10,
    marginBottom: 4,
    color: '#333333',
  },
  section: {
    marginBottom: 20,
  },
  sectionTitle: {
    fontSize: 16,
    fontWeight: 'bold',
    color: palette.primary,
    marginBottom: 12,
    borderBottomWidth: 2,
    borderBottomColor: palette.primary,
    paddingBottom: 4,
  },
  physiologicalRefs: {
    backgroundColor: palette.background,
    padding: 15,
    borderRadius: 8,
    borderLeftWidth: 4,
    borderLeftColor: palette.border,
  },
  refItem: {
    marginBottom: 10,
  },
  refTitle: {
    fontSize: 11,
    fontWeight: 'bold',
    color: palette.primary,
    marginBottom: 2,
  },
  refValue: {
    fontSize: 10,
    color: '#333333',
    marginBottom: 1,
  },
  refNote: {
    fontSize: 8,
    color: '#666666',
    fontStyle: 'italic',
  },
  importantNote: {
    backgroundColor: '#fef3c7',
    padding: 12,
    borderRadius: 8,
    borderLeftWidth: 4,
    borderLeftColor: '#f59e0b',
    marginTop: 15,
  },
  noteTitle: {
    fontSize: 11,
    fontWeight: 'bold',
    color: '#f59e0b',
    marginBottom: 4,
  },
  noteText: {
    fontSize: 9,
    color: '#666666',
    lineHeight: 1.3,
  },
  chartSection: {
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: '#f8f9fa',
    padding: 20,
    borderRadius: 8,
    marginBottom: 20,
    flex: 1,
  },
  chartTitle: {
    fontSize: 16,
    fontWeight: 'bold',
    color: palette.primary,
    textAlign: 'center',
    marginBottom: 15,
  },
  chartImage: {
    width: '100%',
    maxHeight: 500,
    objectFit: 'contain',
  },
  chartPlaceholder: {
    fontSize: 10,
    color: '#666666',
    textAlign: 'center',
    marginBottom: 4,
  },
  sustainedPowersTable: {
    marginTop: 15,
  },
  powerTableHeader: {
    flexDirection: 'row',
    backgroundColor: palette.primary,
    padding: 8,
  },
  powerTableHeaderCell: {
    flex: 1,
    fontSize: 9,
    fontWeight: 'bold',
    color: 'white',
    textAlign: 'center',
  },
  powerTableRow: {
    flexDirection: 'row',
    borderBottomWidth: 1,
    borderBottomColor: '#e5e7eb',
    padding: 6,
    backgroundColor: '#ffffff',
  },
  powerTableCell: {
    flex: 1,
    fontSize: 8,
    textAlign: 'center',
    color: '#333333',
  },
  zonesTable: {
    marginBottom: 15,
  },
  tableHeader: {
    flexDirection: 'row',
    backgroundColor: palette.primary,
    padding: 8,
  },
  tableHeaderCell: {
    fontSize: 9,
    fontWeight: 'bold',
    color: 'white',
    textAlign: 'center',
  },
  zoneCol: { width: '12%' },
  nameCol: { width: '15%' },
  powerCol: { width: '25%' },
  rpeCol: { width: '12%' },
  domainCol: { width: '15%' },
  wkgCol: { width: '21%' },
  tableRow: {
    flexDirection: 'row',
    borderBottomWidth: 1,
    borderBottomColor: '#e5e7eb',
    padding: 6,
    backgroundColor: '#f9fafb',
  },
  tableCell: {
    fontSize: 8,
    textAlign: 'center',
    color: '#333333',
  },
  zoneTag: {
    backgroundColor: palette.primary,
    color: 'white',
    padding: 2,
    borderRadius: 3,
    fontSize: 7,
    fontWeight: 'bold',
  },
  nameTag: {
    padding: 2,
    borderRadius: 3,
    fontSize: 7,
    fontWeight: 'bold',
  },
  domainText: {
    fontSize: 8,
    fontWeight: 'bold',
  },
  domainsLegend: {
    marginTop: 12,
    padding: 10,
    backgroundColor: '#f8f9fa',
    borderRadius: 6,
  },
  legendTitle: {
    fontSize: 12,
    fontWeight: 'bold',
    color: palette.primary,
    marginBottom: 8,
  },
  legendItem: {
    marginBottom: 8,
  },
  legendItemTitle: {
    fontSize: 9,
    fontWeight: 'bold',
    marginBottom: 2,
  },
  legendItemText: {
    fontSize: 8,
    color: '#555555',
    lineHeight: 1.3,
  },
  brandingFooter: {
    marginTop: 'auto',
    alignItems: 'center',
    color: palette.primary,
    fontSize: 10,
    borderTopWidth: 2,
    borderTopColor: palette.primary,
    paddingTop: 15,
  },
  brandingName: {
    fontSize: 12,
    fontWeight: 'bold',
    color: palette.primary,
    marginBottom: 4,
  },
  brandingTitle: {
    fontSize: 10,
    color: '#666666',
    marginBottom: 6,
  },
  brandingContact: {
    fontSize: 9,
    color: '#666666',
    marginBottom: 2,
  },
  brandingWeb: {
    fontSize: 9,
    color: palette.primary,
    fontWeight: 'bold',
  },
});

// CORRECTED: Sustained power calculations using the 3 models in their correct domains
const generateSustainedPowerTable = (models: PowerModels, peso: number) => {
  const durations = [5, 10, 15, 30, 60, 120, 300, 600, 900, 1200, 1800, 3600];
  const results = [];
  
  console.log('=== CORRECTED SUSTAINED POWER CALCULATIONS ===');
  console.log('Models:', { 
    cp: models.cp, 
    wPrime: models.wPrime, 
    apr_params: models.apr_params, 
    pl_params: models.pl_params 
  });
  
  for (const duration of durations) {
    let power = 0;
    let modelUsed = '';
    
    if (duration <= 180) {
      // APR model for durations ≤ 3 minutes (180 seconds)
      power = models.apr_params.po3min + models.apr_params.apr * Math.exp(-models.apr_params.k * duration);
      modelUsed = 'APR';
      console.log(`${modelUsed} (${duration}s): ${power.toFixed(2)}`);
    } else if (duration <= 960) {
      // CP-W' model for 3-16 minutes (180-960 seconds)
      power = models.cp + (models.wPrime / duration);
      modelUsed = 'CP-W\'';
      console.log(`${modelUsed} (${duration}s): ${power.toFixed(2)}`);
    } else {
      // Power Law model for 16 minutes to 1 hour (960-3600 seconds)
      power = models.pl_params.S * Math.pow(duration, models.pl_params.E - 1);
      modelUsed = 'Power Law';
      console.log(`${modelUsed} (${duration}s): ${power.toFixed(2)}`);
    }
    
    const formatTime = (seconds: number) => {
      if (seconds < 60) return `${seconds}s`;
      if (seconds < 3600) {
        const mins = Math.floor(seconds / 60);
        const secs = seconds % 60;
        return secs > 0 ? `${mins}m${secs}s` : `${mins}m`;
      }
      const hours = Math.floor(seconds / 3600);
      const minutes = Math.floor((seconds % 3600) / 60);
      return minutes > 0 ? `${hours}h${minutes}m` : `${hours}h`;
    };
    
    results.push({
      duration: formatTime(duration),
      power: Math.round(power),
      powerPerKg: Math.round((power / peso) * 10) / 10
    });
  }
  
  console.log('=== CORRECTED SUSTAINED POWERS (3 MODELS) ===', results);
  return results;
};

// CORRECTED: VO2max calculation using CP-W' for 5 minutes
const calculateVO2Max = (cp: number, wPrime: number, peso: number): number => {
  // CORRECTED: P5min = CP + (W'/300) perché 300s > 180s (usa CP-W' non APR)
  const p5min = cp + (wPrime / 300);
  console.log('VO2max calculation: P5min =', p5min, 'peso =', peso);
  // Formula: 7.44 × (P5min ÷ Peso) + 27.51
  const vo2max = Math.round((7.44 * (p5min / peso) + 27.51) * 100) / 100;
  console.log('VO2max result:', vo2max);
  return vo2max;
};

// Training zones (unchanged - already correct)
const getTrainingZones = (models: PowerModels, peso: number) => {
  const lt1 = models.lt1_range.estimate;
  const cp = models.cp;
  const p3min = models.apr_params.po3min;

  return [
    {
      zone: 'Z1',
      name: 'EASY',
      color: '#10b981',
      powerRange: `< ${Math.round(lt1 * 0.8)}`,
      powerRangeWkg: `< ${Math.round((lt1 * 0.8) / peso * 10) / 10}`,
      rpe: '1-2',
      domain: 'Moderate',
      domainColor: '#10b981'
    },
    {
      zone: 'Z2',
      name: 'END',
      color: '#22c55e',
      powerRange: `${Math.round(lt1 * 0.8)} - ${Math.round(lt1)}*`,
      powerRangeWkg: `${Math.round((lt1 * 0.8) / peso * 10) / 10} - ${Math.round(lt1 / peso * 10) / 10}*`,
      rpe: '3-4',
      domain: 'Moderate',
      domainColor: '#10b981'
    },
    {
      zone: 'Z3',
      name: 'MED/SST',
      color: '#eab308',
      powerRange: `${Math.round(lt1)} - ${Math.round(cp * 0.92)}`,
      powerRangeWkg: `${Math.round(lt1 / peso * 10) / 10} - ${Math.round((cp * 0.92) / peso * 10) / 10}`,
      rpe: '5-6',
      domain: 'Heavy',
      domainColor: '#eab308'
    },
    {
      zone: 'Z4',
      name: 'CP',
      color: '#f97316',
      powerRange: `${Math.round(cp * 0.92)} - ${Math.round(cp * 1.02)}`,
      powerRangeWkg: `${Math.round((cp * 0.92) / peso * 10) / 10} - ${Math.round((cp * 1.02) / peso * 10) / 10}`,
      rpe: '7',
      domain: 'Heavy',
      domainColor: '#eab308'
    },
    {
      zone: 'Z5',
      name: 'HIIT',
      color: '#ef4444',
      powerRange: `${Math.round(cp * 1.02)} - ${Math.round(p3min)}`,
      powerRangeWkg: `${Math.round((cp * 1.02) / peso * 10) / 10} - ${Math.round(p3min / peso * 10) / 10}`,
      rpe: '>8',
      domain: 'Severe',
      domainColor: '#ef4444'
    },
    {
      zone: 'Z6',
      name: 'SIT',
      color: '#8b5cf6',
      powerRange: `> ${Math.round(p3min)}`,
      powerRangeWkg: `> ${Math.round(p3min / peso * 10) / 10}`,
      rpe: '>8',
      domain: 'Extreme',
      domainColor: '#8b5cf6'
    }
  ];
};

const PdfDocument: React.FC<PdfReportGeneratorProps> = ({ userData, models, chartImage, profileImage, coachData, colorPalette = 'green' }) => {
  const palette = PDF_COLOR_PALETTES[colorPalette] || PDF_COLOR_PALETTES.green;
  const styles = createStyles(palette);
  const sustainedPowerTable = generateSustainedPowerTable(models, userData.peso_kg);
  
  // CORRECTED: Use CP-W' for P5min calculation
  const vo2max = calculateVO2Max(models.cp, models.wPrime, userData.peso_kg);
  
  const trainingZones = getTrainingZones(models, userData.peso_kg);
  
  // Calculate athlete profile
  const athleteProfile = calculateAthleteProfile(models, userData.peso_kg, 'male'); // Default to male, could be parameterized
  
  // CORRECTED: Calculate P50min using Power Law correctly
  const p50min = models.pl_params.S * Math.pow(3000, models.pl_params.E - 1);
  
  const domains = [
    {
      name: 'Moderate',
      color: '#10b981',
      description: 'Mantenendo un\'intensità costante all\'interno di questo dominio, il consumo di O2 rimane costante e stabile, il lattato rimane anch\'esso stabile intorno ai valori basali.'
    },
    {
      name: 'Heavy',
      color: '#eab308',
      description: 'Mantenendo un\'intensità costante all\'interno di questo dominio, il consumo di O2 aumenta inizialmente per poi raggiungere un plateau. Il lattato si assesta a valori superiori rispetto a quelli basali ma tende ad avere un equilibrio tra produzione e smaltimento/riutilizzo.'
    },
    {
      name: 'Severe',
      color: '#ef4444',
      description: 'Mantenendo un\'intensità costante all\'interno di questo dominio, il consumo di O2 tenderà ad aumentare continuamente fino al raggiungimento del 100% del vo2max. Anche il lattato continuerà ad aumentare rompendo l\'equilibrio tra produzione e smaltimento.'
    },
    {
      name: 'Extreme',
      color: '#8b5cf6',
      description: 'Le intensità all\'interno del dominio estremo portano ad esaurimento ancora prima di raggiungere il 100% del consumo di O2.'
    }
  ];

  return (
    <Document>
      {/* PAGINA 1: Informazioni Personali + Riferimenti Fisiologici */}
      <Page size="A4" style={styles.page}>
        <View style={styles.header}>
          <Text style={styles.title}>Report CP - Power Duration Analyzer</Text>
        </View>

        <View style={styles.section}>
          <View style={styles.personalInfo}>
            <Text style={styles.infoTitle}>Informazioni Personali</Text>
            <Text style={styles.infoText}>Data: {new Date().toLocaleDateString('it-IT')}</Text>
            <Text style={styles.infoText}>Nome: {userData.nome} {userData.cognome}</Text>
            <Text style={styles.infoText}>Email: {userData.email}</Text>
            <Text style={styles.infoText}>Peso Corporeo: {userData.peso_kg} kg</Text>
          </View>
        </View>

        <View style={styles.section}>
          <Text style={styles.sectionTitle}>Riferimenti Fisiologici</Text>
          
          <View style={styles.physiologicalRefs}>
            <View style={styles.refItem}>
              <Text style={styles.refTitle}>CP (Critical Power) - Potenza Critica</Text>
              <Text style={styles.refValue}>
                CP: {Math.round(models.cp)} W ({Math.round(models.cp / userData.peso_kg * 10) / 10} W/kg)
              </Text>
              <Text style={styles.refNote}>
                La Potenza Critica (CP) è il marcatore matematico del confine tra il dominio di intensità 'Heavy' e 'Severe'. 
                Fisiologicamente, questo confine non è un singolo punto ma una zona di transizione (spesso definita Maximum Metabolic Steady State - MMSS) 
                in cui il corpo passa da uno stato metabolico stabile a uno instabile. 
                Allenarsi vicino o leggermente sopra la CP porta rapidamente alla perdita dell'equilibrio e all'aumento della fatica.
              </Text>
            </View>

            <View style={styles.refItem}>
              <Text style={styles.refTitle}>LT1 (Prima Soglia Lattacida) - Stima Empirica</Text>
              <Text style={styles.refValue}>
                Range: {Math.round(models.lt1_range.min)} - {Math.round(models.lt1_range.max)} W 
                ({Math.round(models.lt1_range.min / userData.peso_kg * 10) / 10} - {Math.round(models.lt1_range.max / userData.peso_kg * 10) / 10} W/kg)
              </Text>
              <Text style={styles.refValue}>
                Stima: {Math.round(models.lt1_range.estimate)} W ({Math.round(models.lt1_range.estimate / userData.peso_kg * 10) / 10} W/kg)
              </Text>
              <Text style={styles.refNote}>
                La Prima Soglia Lattacida (LT1) rappresenta la fase di transizione tra il dominio 'Moderate' e 'Heavy'. 
                Al di sopra di questa zona, il lattato ematico e il consumo di ossigeno impiegano più tempo per stabilizzarsi. 
                La stima è basata su dati empirici e andrebbe validata con un test specifico.
              </Text>
            </View>

            <View style={styles.refItem}>
              <Text style={styles.refTitle}>W' (W-prime) - Capacità Anaerobica</Text>
              <Text style={styles.refValue}>W': {Math.round(models.wPrime)} J</Text>
              <Text style={styles.refNote}>Energia disponibile sopra la Critical Power</Text>
            </View>

            <View style={styles.refItem}>
              <Text style={styles.refTitle}>Indice di Resistenza</Text>
              <Text style={styles.refValue}>Parametro E: {models.pl_params.E.toFixed(4)}</Text>
              <Text style={styles.refNote}>E è un parametro che descrive quanto lentamente la potenza di un atleta decade nel tempo.</Text>
            </View>

            <View style={styles.refItem}>
              <Text style={styles.refTitle}>VO2max Stimato</Text>
              <Text style={styles.refValue}>{vo2max} ml/kg/min</Text>
              <Text style={styles.refNote}>Calcolato indirettamente Sitko et. al 2022</Text>
            </View>
          </View>

          <View style={styles.importantNote}>
            <Text style={styles.noteTitle}>Nota Importante</Text>
            <Text style={styles.noteText}>
              I valori di Potenza Critica (CP) e delle altre soglie sono stime derivate da modelli matematici e sono soggetti a un margine di errore statistico e fisiologico. 
              Rappresentano la migliore stima di una 'zona di transizione' piuttosto che un valore assoluto.
            </Text>
          </View>
        </View>

        <View style={styles.brandingFooter}>
          {coachData?.isCoach ? (
            <>
              {coachData.coachLogo && (
                <Image 
                  src={coachData.coachLogo} 
                  style={{ width: 40, height: 40, marginBottom: 5, objectFit: 'contain' }} 
                />
              )}
              <Text style={styles.brandingName}>{coachData.coachName}</Text>
              {coachData.coachCompany && (
                <Text style={styles.brandingTitle}>{coachData.coachCompany}</Text>
              )}
              <Text style={styles.brandingContact}>Mail: {coachData.coachEmail}</Text>
            </>
          ) : (
            <>
              <Text style={styles.brandingName}>Michael Pesse</Text>
              <Text style={styles.brandingTitle}>Allenatore Professionista</Text>
              <Text style={styles.brandingContact}>Mail: info@miccoach.it</Text>
              <Text style={styles.brandingWeb}>www.miccoach.it</Text>
            </>
          )}
        </View>
      </Page>

      {/* PAGINA 2: Solo Grafico */}
      <Page size="A4" style={styles.page}>
        <View style={styles.header}>
          <Text style={styles.title}>Curva Potenza-Durata</Text>
        </View>

        <View style={styles.chartSection}>
          {chartImage ? (
            <Image src={chartImage} style={styles.chartImage} />
          ) : (
            <>
              <Text style={styles.chartPlaceholder}>
                Grafico della relazione tra potenza sostenibile e durata dell'esercizio
              </Text>
              <Text style={styles.chartPlaceholder}>
                • Pmax: {Math.round(models.pMax)} W (potenza massima)
              </Text>
              <Text style={styles.chartPlaceholder}>
                • P3min: {Math.round(models.apr_params.po3min)} W (potenza 3 minuti)
              </Text>
              <Text style={styles.chartPlaceholder}>
                • CP: {Math.round(models.cp)} W (critical power)
              </Text>
              <Text style={styles.chartPlaceholder}>
                • LT1: {Math.round(models.lt1_range.estimate)} W (prima soglia)
              </Text>
            </>
          )}
        </View>

        <View style={styles.brandingFooter}>
          {coachData?.isCoach ? (
            <>
              {coachData.coachLogo && (
                <Image 
                  src={coachData.coachLogo} 
                  style={{ width: 40, height: 40, marginBottom: 5, objectFit: 'contain' }} 
                />
              )}
              <Text style={styles.brandingName}>{coachData.coachName}</Text>
              {coachData.coachCompany && (
                <Text style={styles.brandingTitle}>{coachData.coachCompany}</Text>
              )}
              <Text style={styles.brandingContact}>Mail: {coachData.coachEmail}</Text>
            </>
          ) : (
            <>
              <Text style={styles.brandingName}>Michael Pesse</Text>
              <Text style={styles.brandingTitle}>Allenatore Professionista</Text>
              <Text style={styles.brandingContact}>Mail: info@miccoach.it</Text>
              <Text style={styles.brandingWeb}>www.miccoach.it</Text>
            </>
          )}
        </View>
      </Page>

      {/* PAGINA 3: Solo Tabella Potenze Sostenibili */}
      <Page size="A4" style={styles.page}>
        <View style={styles.header}>
          <Text style={styles.title}>Potenze Sostenibili</Text>
        </View>

        <View style={styles.section}>
          <Text style={styles.sectionTitle}>Tabella Potenze Sostenibili (fino a 1 ora)</Text>
          <View style={styles.sustainedPowersTable}>
            <View style={styles.powerTableHeader}>
              <Text style={styles.powerTableHeaderCell}>Durata</Text>
              <Text style={styles.powerTableHeaderCell}>Potenza Assoluta (W)</Text>
              <Text style={styles.powerTableHeaderCell}>Potenza Relativa (W/kg)</Text>
            </View>
            {sustainedPowerTable.map((row, index) => (
              <View key={index} style={styles.powerTableRow}>
                <Text style={styles.powerTableCell}>{row.duration}</Text>
                <Text style={styles.powerTableCell}>{row.power}</Text>
                <Text style={styles.powerTableCell}>{row.powerPerKg}</Text>
              </View>
            ))}
          </View>
        </View>

        <View style={styles.brandingFooter}>
          {coachData?.isCoach ? (
            <>
              {coachData.coachLogo && (
                <Image 
                  src={coachData.coachLogo} 
                  style={{ width: 40, height: 40, marginBottom: 5, objectFit: 'contain' }} 
                />
              )}
              <Text style={styles.brandingName}>{coachData.coachName}</Text>
              {coachData.coachCompany && (
                <Text style={styles.brandingTitle}>{coachData.coachCompany}</Text>
              )}
              <Text style={styles.brandingContact}>Mail: {coachData.coachEmail}</Text>
            </>
          ) : (
            <>
              <Text style={styles.brandingName}>Michael Pesse</Text>
              <Text style={styles.brandingTitle}>Allenatore Professionista</Text>
              <Text style={styles.brandingContact}>Mail: info@miccoach.it</Text>
              <Text style={styles.brandingWeb}>www.miccoach.it</Text>
            </>
          )}
        </View>
      </Page>

      {/* PAGINA 4: Zone di Allenamento e Domini */}
      <Page size="A4" style={styles.page}>
        <View style={styles.header}>
          <Text style={styles.title}>Zone di Allenamento e Domini di Intensità</Text>
        </View>

        <View style={styles.section}>
          <Text style={styles.sectionTitle}>Zone di Potenza Fisiologiche</Text>
          
          <View style={styles.zonesTable}>
            <View style={styles.tableHeader}>
              <Text style={[styles.tableHeaderCell, styles.zoneCol]}>Zona</Text>
              <Text style={[styles.tableHeaderCell, styles.nameCol]}>Nome</Text>
              <Text style={[styles.tableHeaderCell, styles.powerCol]}>Potenza (W)</Text>
              <Text style={[styles.tableHeaderCell, styles.wkgCol]}>Potenza (W/kg)</Text>
              <Text style={[styles.tableHeaderCell, styles.rpeCol]}>RPE</Text>
              <Text style={[styles.tableHeaderCell, styles.domainCol]}>Dominio</Text>
            </View>
            
            {trainingZones.map((zone) => (
              <View key={zone.zone} style={styles.tableRow}>
                <View style={[styles.tableCell, styles.zoneCol, { alignItems: 'center' }]}>
                  <Text style={styles.zoneTag}>{zone.zone}</Text>
                </View>
                <View style={[styles.tableCell, styles.nameCol, { alignItems: 'center' }]}>
                  <Text style={[styles.nameTag, { backgroundColor: zone.color, color: 'white' }]}>
                    {zone.name}
                  </Text>
                </View>
                <Text style={[styles.tableCell, styles.powerCol]}>{zone.powerRange}</Text>
                <Text style={[styles.tableCell, styles.wkgCol]}>{zone.powerRangeWkg}</Text>
                <Text style={[styles.tableCell, styles.rpeCol]}>{zone.rpe}</Text>
                <Text style={[styles.tableCell, styles.domainCol, styles.domainText, { color: zone.domainColor }]}>
                  {zone.domain}
                </Text>
              </View>
            ))}
          </View>

          <View style={styles.domainsLegend}>
            <Text style={styles.legendTitle}>Domini di Intensità</Text>
            {domains.map((domain) => (
              <View key={domain.name} style={styles.legendItem}>
                <Text style={[styles.legendItemTitle, { color: domain.color }]}>
                  {domain.name}:
                </Text>
                <Text style={styles.legendItemText}>{domain.description}</Text>
              </View>
            ))}
            <Text style={[styles.legendItemText, { marginTop: 8, fontSize: 7, fontStyle: 'italic' }]}>
              *Stima basata su dati empirici e parametro E del modello Power Law. 
              Un test del lattato in laboratorio è raccomandato per una determinazione precisa di LT1.
            </Text>
          </View>
        </View>

        <View style={styles.brandingFooter}>
          {coachData?.isCoach ? (
            <>
              {coachData.coachLogo && (
                <Image 
                  src={coachData.coachLogo} 
                  style={{ width: 40, height: 40, marginBottom: 5, objectFit: 'contain' }} 
                />
              )}
              <Text style={styles.brandingName}>{coachData.coachName}</Text>
              {coachData.coachCompany && (
                <Text style={styles.brandingTitle}>{coachData.coachCompany}</Text>
              )}
              <Text style={styles.brandingContact}>Mail: {coachData.coachEmail}</Text>
            </>
          ) : (
            <>
              <Text style={styles.brandingName}>Michael Pesse</Text>
              <Text style={styles.brandingTitle}>Allenatore Professionista</Text>
              <Text style={styles.brandingContact}>Mail: info@miccoach.it</Text>
              <Text style={styles.brandingWeb}>www.miccoach.it</Text>
            </>
          )}
        </View>
      </Page>

      {/* PAGINA 5: Screenshot Profilo Fisiologico */}
      <Page size="A4" style={styles.page}>
        <View style={styles.header}>
          <Text style={styles.title}>Profilo Fisiologico e Allenamento</Text>
        </View>

        <View style={styles.chartSection}>
          {profileImage ? (
            <Image src={profileImage} style={styles.chartImage} />
          ) : (
            <>
              <Text style={styles.chartPlaceholder}>
                Profilo Fisiologico dell'Atleta
              </Text>
              <Text style={styles.chartPlaceholder}>
                Tipologia: {formatProfileType(athleteProfile.profileType)}
              </Text>
              <Text style={styles.chartPlaceholder}>
                Livello: {athleteProfile.rarityTier}
              </Text>
              <Text style={styles.chartPlaceholder}>
                Rating: {athleteProfile.overallRating.toFixed(1)}/10
              </Text>
            </>
          )}
        </View>

        <View style={styles.brandingFooter}>
          {coachData?.isCoach ? (
            <>
              {coachData.coachLogo && (
                <Image 
                  src={coachData.coachLogo} 
                  style={{ width: 40, height: 40, marginBottom: 5, objectFit: 'contain' }} 
                />
              )}
              <Text style={styles.brandingName}>{coachData.coachName}</Text>
              {coachData.coachCompany && (
                <Text style={styles.brandingTitle}>{coachData.coachCompany}</Text>
              )}
              <Text style={styles.brandingContact}>Mail: {coachData.coachEmail}</Text>
            </>
          ) : (
            <>
              <Text style={styles.brandingName}>Michael Pesse</Text>
              <Text style={styles.brandingTitle}>Allenatore Professionista</Text>
              <Text style={styles.brandingContact}>Mail: info@miccoach.it</Text>
              <Text style={styles.brandingWeb}>www.miccoach.it</Text>
            </>
          )}
        </View>
      </Page>
    </Document>
  );
};

export const generatePdfReport = async (userData: UserData, models: PowerModels, chartImage?: string | null, profileImage?: string | null, coachData?: CoachData, colorPalette?: string): Promise<Blob> => {
  try {
    console.log('=== PDF GENERATION START ===');
    console.log('User data:', userData);
    console.log('Models:', models);
    console.log('Chart image available:', !!chartImage);
    console.log('Coach data:', coachData);
    
    const doc = <PdfDocument userData={userData} models={models} chartImage={chartImage} profileImage={profileImage} coachData={coachData} colorPalette={colorPalette} />;
    const pdfBlob = await pdf(doc).toBlob();
    
    console.log('=== PDF GENERATION SUCCESS ===');
    return pdfBlob;
  } catch (error) {
    console.error('=== PDF GENERATION ERROR ===');
    console.error('Error details:', error);
    console.error('Error stack:', error instanceof Error ? error.stack : 'No stack trace');
    
    if (error instanceof Error) {
      console.error('Error name:', error.name);
      console.error('Error message:', error.message);
    }
    
    throw new Error(`Errore nella generazione del PDF: ${error instanceof Error ? error.message : 'Errore sconosciuto'}`);
  }
};

export default PdfDocument;
