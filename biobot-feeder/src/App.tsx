// src/App.tsx

import React, { useState } from 'react';
import { generateVBScript } from './generateScript';
import './App.css'; // This imports your styles

const App = () => {
  // --- STATE MANAGEMENT ---
  const [formData, setFormData] = useState({
    userName: '',
    experimentId: '',
    reactorId: '',
    sampleDelay: 4,
    feedAPump: 'D', feedAFCal: 17.2, feedAFlowRate: 250,
    feedBPump: 'F', feedBFCal: 105, feedBFlowRate: 35,
    glucosePump: 'E', glucoseFCal: 28, glucoseFlowRate: 160
  });

  // Default Schedule Data
  const defaultTimes = [24,48,72,96,120,144,168,192,216,240,264,288,312,336];
  const defaultFeedA = [0,45,0,90,0,120,0,150,0,150,0,0,0,0];
  const defaultFeedB = [0.0,1.5,0.0,3.0,0.0,4.5,0.0,6.0,0.0,7.5,0.0,0.0,0.0,0.0];

  const [schedule, setSchedule] = useState(
    Array.from({ length: 14 }, (_, i) => ({
      id: i + 1,
      time: defaultTimes[i],
      feedA: defaultFeedA[i],
      feedB: defaultFeedB[i]
    }))
  );

  // --- HANDLERS ---
  const handleInputChange = (e: any) => {
    const { id, value } = e.target;
    // Map HTML IDs to State Keys if necessary, or just use name
    setFormData(prev => ({ ...prev, [id]: value }));
  };

  const handleScheduleChange = (index: number, field: string, value: string) => {
    const newSchedule: any = [...schedule];
    newSchedule[index][field] = parseFloat(value);
    setSchedule(newSchedule);
  };

  const add24Hours = (index: number) => {
    if (index === 0) return;
    const prevTime = schedule[index - 1].time;
    const newSchedule = [...schedule];
    newSchedule[index].time = (isNaN(prevTime) ? 0 : prevTime) + 24;
    setSchedule(newSchedule);
  };

  const downloadScript = () => {
    if (!formData.userName || !formData.experimentId || !formData.reactorId) {
      alert('Please complete all required fields (User, Experiment, Reactor).');
      return;
    }
    const scriptContent = generateVBScript({ ...formData, schedule });
    const blob = new Blob([scriptContent], { type: 'text/plain' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${formData.experimentId}_${formData.reactorId}_Autofeed_v5.5.txt`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const PumpSelect = ({ id, value }: any) => (
    <select id={id} value={value} onChange={handleInputChange}>
      {['A', 'B', 'C', 'D', 'E', 'F'].map(p => <option key={p} value={p}>Pump {p}</option>)}
    </select>
  );

  return (
    <div className="container">
      <div className="card">
        <div className="card-header">
          Autofeed Configuration Form v5.5
        </div>

        {/* Experiment Info */}
        <div className="section-title">Experiment Information</div>
        <div className="form-grid">
          <div className="form-field">
            <label htmlFor="userName">User Name:</label>
            <input type="text" id="userName" placeholder="e.g. Colin Larsson" onChange={handleInputChange} />
          </div>
          <div className="form-field">
            <label htmlFor="experimentId">Experiment ID:</label>
            <input type="text" id="experimentId" placeholder="e.g. 706R020725" onChange={handleInputChange} />
          </div>
          <div className="form-field">
            <label htmlFor="reactorId">Reactor ID:</label>
            <input type="text" id="reactorId" placeholder="e.g. H1-4" onChange={handleInputChange} />
          </div>
          <div className="form-field">
            <label htmlFor="sampleDelay">Sample Conf Extension (hrs):</label>
            <input type="number" id="sampleDelay" value={formData.sampleDelay} onChange={handleInputChange} />
            <span className="help-text">After scheduled feed: Script waits X hrs for confirmation...</span>
          </div>
        </div>

        {/* Pump Configuration */}
        <div className="section-title">Pump Configuration</div>
        
        {/* Feed A */}
        <div className="pump-config">
          <div style={{ fontWeight: 600, marginBottom: '10px' }}>Feed A (Major Feed) - Gravimetric</div>
          <div className="pump-config-grid">
            <div className="form-field">
              <label htmlFor="feedAPump">Pump Assignment:</label>
              <PumpSelect id="feedAPump" value={formData.feedAPump} />
              <span className="help-text">Defaults: Dasgip Pump C; SVT Pump D</span>
            </div>
            <div className="form-field">
              <label htmlFor="feedAFCal">F-Calibration:</label>
              <input type="number" id="feedAFCal" value={formData.feedAFCal} onChange={handleInputChange} />
              <span className="help-text">Don't change, unless tubing isn't custom 2mm</span>
            </div>
            <div className="form-field">
              <label htmlFor="feedAFlowRate">Flow Rate (mL/h):</label>
              <input type="number" id="feedAFlowRate" value={formData.feedAFlowRate} onChange={handleInputChange} />
              <span className="help-text">Max ~4200rph / fcal</span>
            </div>
          </div>
        </div>

        {/* Feed B */}
        <div className="pump-config">
          <div style={{ fontWeight: 600, marginBottom: '10px' }}>Feed B - Volumetric</div>
          <div className="pump-config-grid">
            <div className="form-field">
              <label htmlFor="feedBPump">Pump Assignment:</label>
              <PumpSelect id="feedBPump" value={formData.feedBPump} />
              <span className="help-text">Defaults: Dasgip Pump A; SVT Pump F</span>
            </div>
            <div className="form-field">
              <label htmlFor="feedBFCal">F-Calibration:</label>
              <input type="number" id="feedBFCal" value={formData.feedBFCal} onChange={handleInputChange} />
            </div>
            <div className="form-field">
              <label htmlFor="feedBFlowRate">Flow Rate (mL/h):</label>
              <input type="number" id="feedBFlowRate" value={formData.feedBFlowRate} onChange={handleInputChange} />
              <span className="help-text">Max ~4200rph / fcal</span>
            </div>
          </div>
        </div>

        {/* Glucose */}
        <div className="pump-config">
          <div style={{ fontWeight: 600, marginBottom: '10px' }}>Glucose Feed - Gravimetric</div>
          <div className="pump-config-grid">
            <div className="form-field">
              <label htmlFor="glucosePump">Pump Assignment:</label>
              <PumpSelect id="glucosePump" value={formData.glucosePump} />
              <span className="help-text">Defaults: Dasgip Pump D; SVT Pump E</span>
            </div>
            <div className="form-field">
              <label htmlFor="glucoseFCal">F-Calibration:</label>
              <input type="number" id="glucoseFCal" value={formData.glucoseFCal} onChange={handleInputChange} />
            </div>
            <div className="form-field">
              <label htmlFor="glucoseFlowRate">Flow Rate (mL/h):</label>
              <input type="number" id="glucoseFlowRate" value={formData.glucoseFlowRate} onChange={handleInputChange} />
              <span className="help-text">Max ~4200rph / fcal</span>
            </div>
          </div>
        </div>

        {/* Schedule */}
        <div className="section-title">Feed Schedule</div>
        <span className="help-text">Feed event = 1.Feed_B 2.Feed_A 3.Glucose (if requested).
          <br /> Times are in hours after inoculation, leave extra feed times after last feed as is.
          <br /> Save to script repository once generated to access from other stations.</span>
        
        <table>
          <thead>
            <tr>
              <th>Event</th>
              <th>Feed Time (hrs)</th>
              <th>Feed A (g)</th>
              <th>Feed B (mL)</th>
            </tr>
          </thead>
          <tbody>
            {schedule.map((row, index) => (
              <tr key={row.id}>
                <td>Feed #{String(row.id).padStart(2, '0')}</td>
                <td>
                  <input
                    type="number"
                    value={row.time}
                    onChange={(e) => handleScheduleChange(index, 'time', e.target.value)}
                  />
                  {index > 0 && (
                    <button
                      className="plus24-btn"
                      type="button"
                      onClick={() => add24Hours(index)}
                      title="Add 24hrs to previous"
                    >
                      <span>+24</span>
                      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                        <line x1="10" x2="14" y1="2" y2="2" />
                        <line x1="12" x2="15" y1="14" y2="11" />
                        <circle cx="12" cy="14" r="8" />
                      </svg>
                    </button>
                  )}
                </td>
                <td>
                  <input
                    type="number"
                    value={row.feedA}
                    onChange={(e) => handleScheduleChange(index, 'feedA', e.target.value)}
                  />
                </td>
                <td>
                  <input
                    type="number"
                    value={row.feedB}
                    onChange={(e) => handleScheduleChange(index, 'feedB', e.target.value)}
                  />
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        <button type="button" className="submit-btn" onClick={downloadScript}>
          Generate Script
        </button>

        <div className="card-footer">
          Colin Larsson — Autofeed Configuration Form v5.5
        </div>
      </div>
    </div>
  );
};

export default App;