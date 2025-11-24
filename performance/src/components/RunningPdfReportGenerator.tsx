
import React from 'react';
import { pdf, Document, Page, Text, View, StyleSheet, Image } from '@react-pdf/renderer';
import { PdfColorPalette, PDF_COLOR_PALETTES } from '@/utils/pdfColorPalettes';

interface RunningModels {
  valid: boolean;
  cs: number;
  dPrime: number;
  cs_pace: number;
  pl_params: {
    A: number;
    k: number;
  };
  lt1_range: {
    min: number;
    max: number;
    estimate: number;
  };
  vo2max_pace: number;
}

interface CoachData {
  isCoach: boolean;
  coachName?: string;
  coachCompany?: string;
  coachEmail?: string;
  coachLogo?: string;
}

interface UserData {
  nome: string;
  cognome: string;
  email: string;
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
  row: {
    flexDirection: 'row',
    marginBottom: 8,
  },
  label: {
    fontSize: 12,
    fontWeight: 'bold',
    width: '40%',
    color: '#475569',
  },
  value: {
    fontSize: 12,
    width: '60%',
    color: '#1e293b',
  },
  table: {
    width: 'auto',
    borderStyle: 'solid',
    borderWidth: 1,
    borderColor: '#e2e8f0',
  },
  tableRow: {
    margin: 'auto',
    flexDirection: 'row',
  },
  tableCell: {
    borderStyle: 'solid',
    borderWidth: 1,
    borderColor: '#e2e8f0',
    padding: 8,
    fontSize: 10,
    textAlign: 'center',
  },
  tableHeader: {
    backgroundColor: '#f1f5f9',
    fontWeight: 'bold',
    color: '#475569',
  },
  footer: {
    position: 'absolute',
    bottom: 30,
    left: 30,
    right: 30,
    borderTop: '1 solid #e2e8f0',
    paddingTop: 10,
  },
  footerText: {
    fontSize: 10,
    color: '#64748b',
    textAlign: 'center',
    marginBottom: 3,
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
});

const formatPace = (secondsPerKm: number): string => {
  if (!isFinite(secondsPerKm) || secondsPerKm <= 0) return "--:--";
  const minutes = Math.floor(secondsPerKm / 60);
  const seconds = Math.round(secondsPerKm % 60);
  return `${minutes}:${seconds.toString().padStart(2, '0')}`;
};

const formatTime = (totalSeconds: number): string => {
  if (!isFinite(totalSeconds) || totalSeconds <= 0) return "--:--";
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = Math.round(totalSeconds % 60);
  
  if (hours > 0) {
    return `${hours}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
  } else {
    return `${minutes}:${seconds.toString().padStart(2, '0')}`;
  }
};

const calculatePerformancePredictions = (models: RunningModels) => {
  const distances = [
    { name: '1500m', meters: 1500 },
    { name: '3000m', meters: 3000 },
    { name: '5000m', meters: 5000 },
    { name: '10km', meters: 10000 },
    { name: 'Mezza Maratona', meters: 21097 },
    { name: 'Maratona', meters: 42195 },
    { name: '50km', meters: 50000 },
  ];

  return distances.map(distance => {
    // Use the same logic as PerformanceCalculator.tsx
    // Try CS-D' model first - using high precision values
    const timeFromCS = (distance.meters - models.dPrime) / models.cs;
    
    let predictedTime: number;

    if (timeFromCS > 0 && timeFromCS < 1020) { // < 17 minutes
      predictedTime = timeFromCS;
    } else {
      // Use Power Law model: v = S * t^(E-1)  =>  t = (D/S)^(1/E)
      const S = models.pl_params.A;     // velocità base
      const E_minus_1 = models.pl_params.k; // E-1
      const E = E_minus_1 + 1;          // E
      
      // Formula corretta: t = (D/S)^(1/E)
      predictedTime = Math.pow(distance.meters / S, 1 / E);
    }

    // Verifica che il tempo sia valido
    if (!isFinite(predictedTime) || predictedTime <= 0) {
      // Fallback using simple speed calculation
      predictedTime = distance.meters / models.cs;
    }

    // Calcola passo medio
    const avgPace = (predictedTime * 1000) / distance.meters;

    return {
      distance: distance.name,
      time: formatTime(predictedTime),
      pace: formatPace(avgPace),
    };
  });
};

const RunningPdfDocument = ({ userData, models, chartImage, zonesImage, profileImage, coachData, colorPalette = 'green' }: { 
  userData: UserData; 
  models: RunningModels; 
  chartImage?: string | null;
  zonesImage?: string | null;
  profileImage?: string | null;
  coachData?: CoachData;
  colorPalette?: string;
}) => {
  const palette = PDF_COLOR_PALETTES[colorPalette] || PDF_COLOR_PALETTES.green;
  const styles = createStyles(palette);
  const predictions = calculatePerformancePredictions(models);
  const currentDate = new Date().toLocaleDateString('it-IT');

  return (
    <Document>
      {/* Pagina 1: Informazioni e Riferimenti Fisiologici */}
      <Page size="A4" style={styles.page}>
        <View style={styles.header}>
          <Text style={styles.title}>Report CP - Power Duration Analyzer</Text>
        </View>

        <View style={styles.section}>
          <View style={styles.personalInfo}>
            <Text style={styles.infoTitle}>Informazioni Personali</Text>
            <Text style={styles.infoText}>Data: {currentDate}</Text>
            <Text style={styles.infoText}>Nome: {userData.nome} {userData.cognome}</Text>
            <Text style={styles.infoText}>Email: {userData.email}</Text>
          </View>
        </View>

        <View style={styles.section}>
          <Text style={styles.sectionTitle}>Riferimenti Fisiologici</Text>
          
          <View style={styles.physiologicalRefs}>
            <View style={styles.refItem}>
              <Text style={styles.refTitle}>CS (Critical Speed) - Velocità Critica</Text>
              <Text style={styles.refValue}>
                CS: {models.cs.toFixed(3)} m/s ({formatPace(1000 / models.cs)} min/km)
              </Text>
              <Text style={styles.refNote}>
                La Velocità Critica (CS) è il marcatore matematico del confine tra il dominio di intensità 'Heavy' e 'Severe'. 
                Fisiologicamente, questo confine non è un singolo punto ma una zona di transizione (spesso definita Maximum Metabolic Steady State - MMSS) 
                in cui il corpo passa da uno stato metabolico stabile a uno instabile. 
                Correre vicino o leggermente sopra la CS porta rapidamente alla perdita dell'equilibrio e all'aumento della fatica.
              </Text>
            </View>

            <View style={styles.refItem}>
              <Text style={styles.refTitle}>LT1 (Prima Soglia Lattacida) - Stima Empirica</Text>
              <Text style={styles.refValue}>
                Range: {formatPace(models.lt1_range.min)} - {formatPace(models.lt1_range.max)} min/km
              </Text>
              <Text style={styles.refValue}>
                Stima: {formatPace(models.lt1_range.estimate)} min/km
              </Text>
              <Text style={styles.refNote}>
                La Prima Soglia Lattacida (LT1) rappresenta la fase di transizione tra il dominio 'Moderate' e 'Heavy'. 
                Al di sopra di questa zona, il lattato ematico e il consumo di ossigeno impiegano più tempo per stabilizzarsi. 
                La stima è basata su dati empirici e andrebbe validata con un test specifico.
              </Text>
            </View>

            <View style={styles.refItem}>
              <Text style={styles.refTitle}>D' (D-prime) - Capacità Anaerobica</Text>
              <Text style={styles.refValue}>D': {Math.round(models.dPrime)} m</Text>
              <Text style={styles.refNote}>Distanza disponibile sopra la Critical Speed</Text>
            </View>

            <View style={styles.refItem}>
              <Text style={styles.refTitle}>Indice di Resistenza</Text>
              <Text style={styles.refValue}>Parametro E: {(models.pl_params.k + 1).toFixed(4)}</Text>
              <Text style={styles.refNote}>E è un parametro che descrive quanto lentamente la velocità di un atleta decade nel tempo.</Text>
            </View>

            <View style={styles.refItem}>
              <Text style={styles.refTitle}>MAP (3min pace)</Text>
              <Text style={styles.refValue}>{formatPace(models.vo2max_pace)} min/km</Text>
              <Text style={styles.refNote}>Passo teorico per 3 minuti di corsa</Text>
            </View>
          </View>

          <View style={styles.importantNote}>
            <Text style={styles.noteTitle}>Nota Importante</Text>
            <Text style={styles.noteText}>
              I valori di Velocità Critica (CS) e delle altre soglie sono stime derivate da modelli matematici e sono soggetti a un margine di errore statistico e fisiologico. 
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
          <Text style={styles.title}>Curva Velocità-Durata</Text>
        </View>

        <View style={styles.chartSection}>
          {chartImage ? (
            <Image src={chartImage} style={styles.chartImage} />
          ) : (
            <>
              <Text style={styles.chartPlaceholder}>
                Grafico della relazione tra velocità sostenibile e durata dell'esercizio
              </Text>
              <Text style={styles.chartPlaceholder}>
                • Velocità Massima: {models.cs.toFixed(3)} m/s
              </Text>
              <Text style={styles.chartPlaceholder}>
                • CS: {models.cs.toFixed(3)} m/s (velocità critica)
              </Text>
              <Text style={styles.chartPlaceholder}>
                • LT1: {formatPace(models.lt1_range.estimate)} min/km (prima soglia)
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

      {/* PAGINA 3: Solo Tabella Previsioni Performance */}
      <Page size="A4" style={styles.page}>
        <View style={styles.header}>
          <Text style={styles.title}>Previsioni Performance</Text>
        </View>

        <View style={styles.section}>
          <Text style={styles.sectionTitle}>Tabella Previsioni Performance</Text>
          <View style={styles.table}>
            <View style={styles.tableRow}>
              <Text style={[styles.tableCell, styles.tableHeader, { width: '40%' }]}>
                Distanza
              </Text>
              <Text style={[styles.tableCell, styles.tableHeader, { width: '30%' }]}>
                Tempo Previsto
              </Text>
              <Text style={[styles.tableCell, styles.tableHeader, { width: '30%' }]}>
                Passo Medio
              </Text>
            </View>
            
            {predictions.map((pred, index) => (
              <View key={index} style={styles.tableRow}>
                <Text style={[styles.tableCell, { width: '40%' }]}>{pred.distance}</Text>
                <Text style={[styles.tableCell, { width: '30%' }]}>{pred.time}</Text>
                <Text style={[styles.tableCell, { width: '30%' }]}>{pred.pace}</Text>
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

      {/* PAGINA 4: Zone di Allenamento */}
      <Page size="A4" style={styles.page}>
        <View style={styles.header}>
          <Text style={styles.title}>Zone di Allenamento</Text>
        </View>

        <View style={styles.chartSection}>
          {zonesImage ? (
            <Image src={zonesImage} style={styles.chartImage} />
          ) : (
            <Text style={styles.chartPlaceholder}>
              Tabella delle zone di allenamento fisiologiche
            </Text>
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

      {/* PAGINA 5: Screenshot Profilo Fisiologico Running */}
      <Page size="A4" style={styles.page}>
        <View style={styles.header}>
          <Text style={styles.title}>Profilo Fisiologico e Allenamento</Text>
        </View>

        <View style={styles.chartSection}>
          {profileImage ? (
            <Image src={profileImage} style={styles.chartImage} />
          ) : (
            <Text style={styles.chartPlaceholder}>
              Profilo Fisiologico dell'Atleta Running
            </Text>
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

export const generateRunningPdfReport = async (
  userData: UserData,
  models: RunningModels,
  chartImage?: string | null,
  zonesImage?: string | null,
  profileImage?: string | null,
  coachData?: CoachData,
  colorPalette?: string
): Promise<Blob> => {
  const doc = <RunningPdfDocument userData={userData} models={models} chartImage={chartImage} zonesImage={zonesImage} profileImage={profileImage} coachData={coachData} colorPalette={colorPalette} />;
  const pdfBlob = await pdf(doc).toBlob();
  return pdfBlob;
};
