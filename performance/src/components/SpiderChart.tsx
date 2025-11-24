
import React from 'react';
import { AthleteStats } from '@/utils/athleteCardCalculations';
import { RunningAthleteStats } from '@/utils/runningAthleteCalculations';
import { ProfileTheme } from '@/utils/profileThemes';
import { RarityLevel } from '@/utils/raritySystem';

interface SpiderChartProps {
  stats: AthleteStats | RunningAthleteStats;
  theme: ProfileTheme;
  rarityLevel: RarityLevel;
  isRunning?: boolean;
  profileType?: string;
  profileName?: string;
}

const SpiderChart: React.FC<SpiderChartProps> = ({ 
  stats, 
  theme, 
  rarityLevel,
  isRunning = false,
  profileType,
  profileName
}) => {
  const size = 500;
  const center = size / 2;
  const maxRadius = 120;
  
  // Get stat names and values
  const getStatLabels = () => {
    if (isRunning) {
      return ['1500M', '5000M', '10KM', 'MEZZA', 'MARA'];
    }
    return ['PMAX', 'C. ANAEROBICA', 'P. AEROBICA MAX', 'POT. CRITICA', 'ENDURANCE'];
  };

  const statLabels = getStatLabels();
  const statValues = Object.values(stats);
  const maxValue = 100;

  // Calculate pentagon vertices (5 points, starting from top)
  const getVertexPosition = (index: number, radius: number) => {
    const angle = (index * 2 * Math.PI) / 5 - Math.PI / 2; // Start from top
    return {
      x: center + Math.cos(angle) * radius,
      y: center + Math.sin(angle) * radius
    };
  };

  // Generate grid lines (concentric pentagons)
  const generateGridLines = () => {
    const gridLevels = [20, 40, 60, 80, 100];
    return gridLevels.map((level, levelIndex) => {
      const radius = (level / maxValue) * maxRadius;
      const points = Array.from({ length: 5 }, (_, i) => {
        const pos = getVertexPosition(i, radius);
        return `${pos.x},${pos.y}`;
      }).join(' ');
      
      return (
        <polygon
          key={levelIndex}
          points={points}
          fill="none"
          stroke="rgba(189, 189, 189, 0.2)"
          strokeWidth="1"
        />
      );
    });
  };

  // Generate axis lines from center to vertices
  const generateAxisLines = () => {
    return Array.from({ length: 5 }, (_, i) => {
      const vertex = getVertexPosition(i, maxRadius);
      return (
        <line
          key={i}
          x1={center}
          y1={center}
          x2={vertex.x}
          y2={vertex.y}
          stroke="rgba(189, 189, 189, 0.2)"
          strokeWidth="1"
        />
      );
    });
  };

  // Generate athlete's performance shape with rarity effects
  const generateAthleteShape = () => {
    const points = statValues.map((value, index) => {
      const normalizedValue = Math.min(value, maxValue);
      const radius = (normalizedValue / maxValue) * maxRadius;
      const pos = getVertexPosition(index, radius);
      return `${pos.x},${pos.y}`;
    }).join(' ');

    const fillOpacity = 0.3;
    
    // Rarity-based glow effects
    const hasGlow = rarityLevel === 'LEGGENDA' || rarityLevel === 'GOAT';
    const glowRadius = rarityLevel === 'GOAT' ? '12' : '8';
    const glowOpacity = rarityLevel === 'GOAT' ? '0.8' : '0.6';
    const filterId = `neon-glow-${rarityLevel}-${Math.random().toString(36).substr(2, 9)}`;

    return (
      <g>
        {/* SVG Filter Definition for Neon Glow Effect */}
        {hasGlow && (
          <defs>
            <filter id={filterId} x="-50%" y="-50%" width="200%" height="200%">
              {/* Multiple blur layers for neon effect */}
              <feGaussianBlur in="SourceGraphic" stdDeviation={glowRadius} result="coloredBlur"/>
              <feGaussianBlur in="SourceGraphic" stdDeviation={parseInt(glowRadius) * 0.5} result="innerGlow"/>
              <feGaussianBlur in="SourceGraphic" stdDeviation={parseInt(glowRadius) * 2} result="outerGlow"/>
              
              {/* Merge layers for complex glow */}
              <feMerge>
                <feMergeNode in="outerGlow" opacity={glowOpacity}/>
                <feMergeNode in="coloredBlur" opacity={glowOpacity}/>
                <feMergeNode in="innerGlow" opacity="0.9"/>
                <feMergeNode in="SourceGraphic"/>
              </feMerge>
            </filter>
          </defs>
        )}

        {/* Fill shape */}
        <polygon
          points={points}
          fill={theme.primaryColor}
          fillOpacity={fillOpacity}
        />
        
        {/* Border with neon glow effect for LEGGENDA+ */}
        <polygon
          points={points}
          fill="none"
          stroke={theme.primaryColor}
          strokeWidth={hasGlow ? "4" : "3"}
          strokeOpacity={hasGlow ? "1" : "0.9"}
          filter={hasGlow ? `url(#${filterId})` : 'none'}
        />
      </g>
    );
  };

  // Generate labels with two-line layout (Value + Metric name)
  const generateLabels = () => {
    return statLabels.map((label, index) => {
      const labelRadius = maxRadius + 60;
      const pos = getVertexPosition(index, labelRadius);
      const value = statValues[index];
      
      return (
        <g key={index}>
          {/* VALUE - First line (most prominent) */}
          <text
            x={pos.x}
            y={pos.y - 8}
            textAnchor="middle"
            className="font-bold"
            style={{
              fontSize: '28px',
              fill: '#FFFFFF',
              fontFamily: 'Inter, sans-serif',
              fontWeight: 700
            }}
          >
            {value}
          </text>
          
          {/* METRIC NAME - Second line (supporting element) */}
          <text
            x={pos.x}
            y={pos.y + 16}
            textAnchor="middle"
            className="font-normal uppercase tracking-wider"
            style={{
              fontSize: '14px',
              fill: '#BDBDBD',
              fontFamily: 'Inter, sans-serif'
            }}
          >
            {label}
          </text>
        </g>
      );
    });
  };

  return (
    <div className="relative flex justify-center items-center" style={{ width: '100%', height: '450px', marginTop: '20px' }}>
      <svg
        width={size}
        height={size}
        viewBox={`0 0 ${size} ${size}`}
        className="overflow-visible"
      >
        {/* Grid lines */}
        {generateGridLines()}
        
        {/* Axis lines */}
        {generateAxisLines()}
        
        {/* Athlete performance shape */}
        {generateAthleteShape()}
        
        {/* Labels with two-line layout */}
        {generateLabels()}
      </svg>
      
    </div>
  );
};

export default SpiderChart;
