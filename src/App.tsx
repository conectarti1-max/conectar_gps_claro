/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import L from 'leaflet';
import {
  Building2,
  Upload,
  Download,
  Search,
  Filter,
  Plus,
  RefreshCw,
  MapPin,
  FileSpreadsheet,
  BarChart3,
  ListFilter,
  SlidersHorizontal,
  Edit3,
  Trash2,
  Eye,
  CheckCircle2,
  AlertTriangle,
  Clock,
  XCircle,
  HelpCircle,
  Sparkles,
  ChevronRight,
  ChevronLeft,
  Layers,
  Globe,
  FileText,
  Check,
  X,
  Navigation,
  Compass,
  Route,
  Target
} from 'lucide-react';

// Interface for POP FLM Location record
export interface PopLocation {
  id: string;
  siteId: string;
  name: string;
  region: string;
  province?: string;
  district?: string;
  status: string;
  technology: string;
  address: string;
  ipVlan?: string;
  coordinator?: string;
  observations?: string;
  lat: number;
  lng: number;
  dateAdded?: string;
  [key: string]: any;
}

// Sample initial data simulating the POP_FLM_Ago_2026 sheet with accurate coordinates in Peru
const SAMPLE_POP_DATA: PopLocation[] = [
  {
    id: '1',
    siteId: 'POP-LIM-01',
    name: 'Central Miraflores FLM',
    region: 'Lima',
    province: 'Lima',
    district: 'Miraflores',
    status: 'Operativo',
    technology: 'FTTH / GPON',
    address: 'Av. Larco 1230',
    ipVlan: '10.200.10.1',
    coordinator: 'Carlos Mendoza',
    observations: 'Nodo principal sector sur costero. Respaldado con grupo electrógeno.',
    lat: -12.1218,
    lng: -77.0297,
    dateAdded: '2026-08-01'
  },
  {
    id: '2',
    siteId: 'POP-LIM-02',
    name: 'San Isidro Financial Core',
    region: 'Lima',
    province: 'Lima',
    district: 'San Isidro',
    status: 'Operativo',
    technology: 'FTTH / 10G',
    address: 'Calle Los Pinos 450',
    ipVlan: '10.200.12.1',
    coordinator: 'Ana Lucía Torres',
    observations: 'Alta concentración corporativa. Redundancia de fibra óptica.',
    lat: -12.0975,
    lng: -77.0365,
    dateAdded: '2026-08-02'
  },
  {
    id: '3',
    siteId: 'POP-AQP-01',
    name: 'Arequipa Centro Mayor',
    region: 'Arequipa',
    province: 'Arequipa',
    district: 'Cercado',
    status: 'En Mantenimiento',
    technology: 'Radioenlace / FTTH',
    address: 'Av. Ejército 710',
    ipVlan: '10.210.1.5',
    coordinator: 'Roberto Gómez',
    observations: 'Mantenimiento preventivo de baterías programado.',
    lat: -16.3989,
    lng: -71.5350,
    dateAdded: '2026-08-03'
  },
  {
    id: '4',
    siteId: 'POP-TRU-01',
    name: 'Trujillo Norte Hub',
    region: 'La Libertad',
    province: 'Trujillo',
    district: 'Trujillo',
    status: 'Operativo',
    technology: 'FTTH',
    address: 'Calle España 205',
    ipVlan: '10.220.5.10',
    coordinator: 'María Elena Silva',
    observations: 'Funcionamiento óptimo. Ampliación de puertos GPON completada.',
    lat: -8.1116,
    lng: -79.0287,
    dateAdded: '2026-08-04'
  },
  {
    id: '5',
    siteId: 'POP-CUS-01',
    name: 'Cusco Imperial Core',
    region: 'Cusco',
    province: 'Cusco',
    district: 'Wanchaq',
    status: 'Planificado',
    technology: 'FTTH / HFC',
    address: 'Av. de la Cultura 1520',
    ipVlan: '10.230.2.1',
    coordinator: 'Luis Huamán',
    observations: 'Pendiente permisos municipales para tendido de fibra troncal.',
    lat: -13.5319,
    lng: -71.9675,
    dateAdded: '2026-08-05'
  },
  {
    id: '6',
    siteId: 'POP-PIU-01',
    name: 'Piura Sol del Norte',
    region: 'Piura',
    province: 'Piura',
    district: 'Piura',
    status: 'Operativo',
    technology: 'FTTH',
    address: 'Av. Grau 890',
    ipVlan: '10.240.8.1',
    coordinator: 'Jorge Zapata',
    observations: 'Revisión climática superada con éxito.',
    lat: -5.1945,
    lng: -80.6328,
    dateAdded: '2026-08-06'
  },
  {
    id: '7',
    siteId: 'POP-HYO-01',
    name: 'Huancayo Centro Andino',
    region: 'Junín',
    province: 'Huancayo',
    district: 'El Tambo',
    status: 'Alerta',
    technology: 'Radioenlace',
    address: 'Av. Giráldez 410',
    ipVlan: '10.250.3.4',
    coordinator: 'Rosa Quispe',
    observations: 'Fluctuación de energía comercial detectada. Usando UPS de respaldo.',
    lat: -12.0651,
    lng: -75.2049,
    dateAdded: '2026-08-07'
  },
  {
    id: '8',
    siteId: 'POP-ICA-01',
    name: 'Ica Oasis Hub',
    region: 'Ica',
    province: 'Ica',
    district: 'Ica',
    status: 'Operativo',
    technology: 'FTTH',
    address: 'Calle Cajamarca 120',
    ipVlan: '10.260.1.2',
    coordinator: 'Pedro Benavides',
    observations: 'Nodo operativo estable.',
    lat: -14.0678,
    lng: -75.7286,
    dateAdded: '2026-08-08'
  }
];

export default function App() {
  const [locations, setLocations] = useState<PopLocation[]>(SAMPLE_POP_DATA);
  const [activeTab, setActiveTab] = useState<'table' | 'dashboard' | 'map'>('table');
  const [fileName, setFileName] = useState<string>('Site_POP_FLM_Ago_2026 ver 2.0 _12 AGOSTO.xlsx (Hoja: POP_FLM_Ago_2026)');
  const [sheetNames, setSheetNames] = useState<string[]>(['POP_FLM_Ago_2026']);
  const [selectedSheet, setSelectedSheet] = useState<string>('POP_FLM_Ago_2026');
  const [rawWorkbook, setRawWorkbook] = useState<XLSX.WorkBook | null>(null);

  // Filters & Search
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedRegion, setSelectedRegion] = useState('ALL');
  const [selectedStatus, setSelectedStatus] = useState('ALL');
  const [selectedTech, setSelectedTech] = useState('ALL');

  // Modal State
  const [modalMode, setModalMode] = useState<'view' | 'edit' | 'create' | null>(null);
  const [currentLocation, setCurrentLocation] = useState<PopLocation | null>(null);

  // User GPS & In-App Navigation State
  const [userCoords, setUserCoords] = useState<{ lat: number; lng: number }>({ lat: -12.0464, lng: -77.0428 });
  const [gpsAcquired, setGpsAcquired] = useState(false);
  const [navActive, setNavActive] = useState(false);
  const [targetLocation, setTargetLocation] = useState<PopLocation | null>(null);
  const [routeInfo, setRouteInfo] = useState<{ distance: string; duration: string; steps: string[] } | null>(null);
  const [navLoading, setNavLoading] = useState(false);

  // AI Insights State
  const [aiModalOpen, setAiModalOpen] = useState(false);
  const [aiLoading, setAiLoading] = useState(false);
  const [aiReport, setAiReport] = useState<string | null>(null);

  // Pagination
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 6;

  const fileInputRef = useRef<HTMLInputElement>(null);
  const mapRef = useRef<HTMLDivElement>(null);
  const leafletMapInstance = useRef<L.Map | null>(null);
  const markersLayerRef = useRef<L.LayerGroup | null>(null);
  const routeLayerRef = useRef<L.LayerGroup | null>(null);

  // Automatically request user GPS location on app load
  useEffect(() => {
    if (navigator.geolocation) {
      navigator.geolocation.getCurrentPosition(
        (pos) => {
          setUserCoords({ lat: pos.coords.latitude, lng: pos.coords.longitude });
          setGpsAcquired(true);
        },
        (err) => {
          console.warn('GPS geolocation permission denied or unavailable:', err);
        },
        { enableHighAccuracy: true, timeout: 10000 }
      );
    }
  }, []);

  // Handle Excel File Upload
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        setRawWorkbook(wb);
        const sheets = wb.SheetNames;
        setSheetNames(sheets);

        let targetSheet = sheets.find(s => s.toLowerCase().includes('pop')) || sheets[0];
        setSelectedSheet(targetSheet);

        parseSheetData(wb, targetSheet);
      } catch (err) {
        console.error('Error parsing excel:', err);
        alert('Error al leer el archivo Excel.');
      }
    };
    reader.readAsBinaryString(file);
  };

  const parseSheetData = (wb: XLSX.WorkBook, sheetName: string) => {
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json<any>(ws, { header: 1 });

    if (data.length < 2) {
      alert('La hoja seleccionada está vacía.');
      return;
    }

    let headerRowIdx = 0;
    for (let i = 0; i < Math.min(5, data.length); i++) {
      const row = data[i];
      if (row && row.some((cell: any) => typeof cell === 'string' && (cell.toLowerCase().includes('site') || cell.toLowerCase().includes('sitio') || cell.toLowerCase().includes('pop') || cell.toLowerCase().includes('nombre') || cell.toLowerCase().includes('region')))) {
        headerRowIdx = i;
        break;
      }
    }

    const headers: string[] = data[headerRowIdx].map((h: any, idx: number) => h ? String(h).trim() : `Col_${idx}`);
    const rows = data.slice(headerRowIdx + 1);

    const parsedLocations: PopLocation[] = rows.filter(row => row && row.some(cell => cell !== undefined && cell !== '')).map((row, index) => {
      const record: any = {
        id: `row-${index}`,
        siteId: `POP-${index + 100}`,
        name: `Sitio ${index + 1}`,
        status: 'Operativo',
        technology: 'FTTH',
        region: 'Lima',
        address: 'S/N',
        lat: -12.0464 + (Math.random() - 0.5) * 0.5,
        lng: -77.0428 + (Math.random() - 0.5) * 0.5
      };
      headers.forEach((header, idx) => {
        const val = row[idx];
        if (val !== undefined) {
          const lowerH = header.toLowerCase();
          if (lowerH.includes('id') || lowerH.includes('codigo') || lowerH.includes('sitio_id')) record.siteId = String(val);
          else if (lowerH.includes('nombre') || lowerH.includes('site') || lowerH.includes('estacion') || lowerH.includes('pop')) record.name = String(val);
          else if (lowerH.includes('region') || lowerH.includes('departamento')) record.region = String(val);
          else if (lowerH.includes('provincia')) record.province = String(val);
          else if (lowerH.includes('distrito')) record.district = String(val);
          else if (lowerH.includes('estado') || lowerH.includes('status')) record.status = String(val);
          else if (lowerH.includes('tecnolog') || lowerH.includes('tech')) record.technology = String(val);
          else if (lowerH.includes('direccion') || lowerH.includes('ubicacion') || lowerH.includes('address')) record.address = String(val);
          else if (lowerH.includes('coordinador') || lowerH.includes('responsable')) record.coordinator = String(val);
          else if (lowerH.includes('observacion') || lowerH.includes('comentario') || lowerH.includes('nota')) record.observations = String(val);
          else if (lowerH.includes('lat')) record.lat = parseFloat(val) || record.lat;
          else if (lowerH.includes('lng') || lowerH.includes('lon')) record.lng = parseFloat(val) || record.lng;
          else record[header] = val;
        }
      });
      return record as PopLocation;
    });

    if (parsedLocations.length > 0) {
      setLocations(parsedLocations);
    }
  };

  const handleSheetChange = (sheetName: string) => {
    setSelectedSheet(sheetName);
    if (rawWorkbook) {
      parseSheetData(rawWorkbook, sheetName);
    }
  };

  const handleResetDemo = () => {
    setLocations(SAMPLE_POP_DATA);
    setFileName('Site_POP_FLM_Ago_2026 ver 2.0 _12 AGOSTO.xlsx (Hoja: POP_FLM_Ago_2026)');
    setSheetNames(['POP_FLM_Ago_2026']);
    setSelectedSheet('POP_FLM_Ago_2026');
    setRawWorkbook(null);
  };

  const handleExportExcel = () => {
    const ws = XLSX.utils.json_to_sheet(locations);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, selectedSheet || 'POP_FLM_Ago_2026');
    XLSX.writeFile(wb, `POP_FLM_Gestion_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Filtered Locations
  const filteredLocations = useMemo(() => {
    return locations.filter(loc => {
      const matchSearch = searchTerm === '' ||
        Object.values(loc).some(val => String(val).toLowerCase().includes(searchTerm.toLowerCase()));

      const matchRegion = selectedRegion === 'ALL' || loc.region === selectedRegion;
      const matchStatus = selectedStatus === 'ALL' || loc.status === selectedStatus;
      const matchTech = selectedTech === 'ALL' || loc.technology === selectedTech;

      return matchSearch && matchRegion && matchStatus && matchTech;
    });
  }, [locations, searchTerm, selectedRegion, selectedStatus, selectedTech]);

  const regions = useMemo(() => Array.from(new Set(locations.map(l => l.region).filter(Boolean))), [locations]);
  const statuses = useMemo(() => Array.from(new Set(locations.map(l => l.status).filter(Boolean))), [locations]);
  const technologies = useMemo(() => Array.from(new Set(locations.map(l => l.technology).filter(Boolean))), [locations]);

  const totalPages = Math.ceil(filteredLocations.length / itemsPerPage) || 1;
  const paginatedLocations = useMemo(() => {
    const start = (currentPage - 1) * itemsPerPage;
    return filteredLocations.slice(start, start + itemsPerPage);
  }, [filteredLocations, currentPage]);

  // When clicking on an address / location row -> Jump to map, center on site with logo/marker popup
  const handleSelectLocationOnMap = (loc: PopLocation) => {
    setActiveTab('map');
    setTargetLocation(loc);
    setNavActive(false);

    if (leafletMapInstance.current) {
      try {
        leafletMapInstance.current.setView([loc.lat, loc.lng], 16, { animate: false });
      } catch (err) {
        console.error('setView error:', err);
      }
    }
  };

  // Start In-App Navigation Route Calculation from user GPS to selected site
  const startInAppNavigation = (loc: PopLocation) => {
    setTargetLocation(loc);
    setActiveTab('map');
    setNavActive(true);
    setNavLoading(true);

    calculateRoute(userCoords, { lat: loc.lat, lng: loc.lng });
  };

  const calculateRoute = async (start: { lat: number; lng: number }, end: { lat: number; lng: number }) => {
    try {
      const res = await fetch(`https://router.project-osrm.org/route/v1/driving/${start.lng},${start.lat};${end.lng},${end.lat}?overview=full&geometries=geojson&steps=true`);
      const data = await res.json();

      if (data && data.routes && data.routes.length > 0) {
        const route = data.routes[0];
        const distanceKm = (route.distance / 1000).toFixed(1) + ' km';
        const durationMin = Math.round(route.duration / 60) + ' min';

        let stepsList: string[] = [];
        if (route.legs && route.legs[0] && route.legs[0].steps) {
          stepsList = route.legs[0].steps.map((s: any) => s.maneuver.instruction || 'Continuar por la vía principal');
        } else {
          stepsList = ['Dirigirse hacia el destino por la vía principal', 'Llegada al nodo POP FLM'];
        }

        setRouteInfo({
          distance: distanceKm,
          duration: durationMin,
          steps: stepsList
        });

        drawRouteOnMap(route.geometry.coordinates, start, end);
      } else {
        fallbackRouteDraw(start, end);
      }
    } catch (e) {
      console.warn('OSRM routing fetch failed:', e);
      fallbackRouteDraw(start, end);
    } finally {
      setNavLoading(false);
    }
  };

  const fallbackRouteDraw = (start: { lat: number; lng: number }, end: { lat: number; lng: number }) => {
    const R = 6371;
    const dLat = (end.lat - start.lat) * (Math.PI / 180);
    const dLon = (end.lng - start.lng) * (Math.PI / 180);
    const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
        Math.cos(start.lat * (Math.PI / 180)) * Math.cos(end.lat * (Math.PI / 180)) *
        Math.sin(dLon / 2) * Math.sin(dLon / 2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    const dist = (R * c).toFixed(1) + ' km';

    setRouteInfo({
      distance: dist,
      duration: '~' + Math.round(parseFloat(dist) * 2.5) + ' min',
      steps: [
        `Partir desde su posición actual GPS (${start.lat.toFixed(4)}, ${start.lng.toFixed(4)})`,
        'Seguir ruta directa hacia el nodo POP',
        'Llegada al destino final'
      ]
    });

    drawRouteOnMap([[start.lng, start.lat], [end.lng, end.lat]], start, end);
    setNavLoading(false);
  };

  const drawRouteOnMap = (coords: [number, number][], start: { lat: number; lng: number }, end: { lat: number; lng: number }) => {
    if (!leafletMapInstance.current) return;
    const map = leafletMapInstance.current;

    try {
      if (!routeLayerRef.current) {
        routeLayerRef.current = L.layerGroup().addTo(map);
      } else {
        routeLayerRef.current.clearLayers();
      }

      const latLngs = coords.map(c => [c[1], c[0]] as [number, number]);

      const polyline = L.polyline(latLngs, {
        color: '#4f46e5',
        weight: 5,
        opacity: 0.85,
        dashArray: '8, 8'
      });
      routeLayerRef.current.addLayer(polyline);

      const userIcon = L.divIcon({
        className: 'user-gps-icon',
        html: `<div style="background-color: #3b82f6; width: 26px; height: 26px; border-radius: 50%; border: 3px solid white; box-shadow: 0 0 12px rgba(59,130,246,0.9); display: flex; align-items: center; justify-content: center; color: white; font-size: 11px;">📍</div>`,
        iconSize: [26, 26],
        iconAnchor: [13, 13]
      });
      const userMarker = L.marker([start.lat, start.lng], { icon: userIcon }).bindPopup('<b>Su ubicación GPS actual</b>');
      routeLayerRef.current.addLayer(userMarker);

      map.fitBounds(polyline.getBounds(), { padding: [60, 60], animate: false });
    } catch (err) {
      console.error('drawRouteOnMap error:', err);
    }
  };

  // Initialize Leaflet Map cleanly when activeTab === 'map'
  useEffect(() => {
    if (activeTab === 'map') {
      const timer = setTimeout(() => {
        if (!mapRef.current) return;

        if (leafletMapInstance.current) {
          try {
            leafletMapInstance.current.remove();
          } catch (e) {
            console.error('Map cleanup error:', e);
          }
          leafletMapInstance.current = null;
        }

        const map = L.map(mapRef.current, {
          zoomControl: true,
          fadeAnimation: false,
          zoomAnimation: false
        }).setView([userCoords.lat, userCoords.lng], 13);

        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
          attribution: '&copy; OpenStreetMap contributors'
        }).addTo(map);

        const markersLayer = L.layerGroup().addTo(map);
        markersLayerRef.current = markersLayer;

        const routeLayer = L.layerGroup().addTo(map);
        routeLayerRef.current = routeLayer;

        leafletMapInstance.current = map;

        // Add user GPS marker
        const userIcon = L.divIcon({
          className: 'user-gps-icon',
          html: `<div style="background-color: #3b82f6; width: 24px; height: 24px; border-radius: 50%; border: 3px solid white; box-shadow: 0 0 10px rgba(59,130,246,0.8); display: flex; align-items: center; justify-content: center; color: white; font-size: 10px;">📍</div>`,
          iconSize: [24, 24],
          iconAnchor: [12, 12]
        });
        L.marker([userCoords.lat, userCoords.lng], { icon: userIcon }).bindPopup('<b>Usted está aquí (GPS)</b>').addTo(markersLayer);

        // Add POP location markers with site logo/icon
        filteredLocations.forEach(loc => {
          if (loc.lat && loc.lng) {
            const statusColor =
              loc.status?.toLowerCase().includes('operativo') ? '#10b981' :
              loc.status?.toLowerCase().includes('mantenimiento') ? '#f59e0b' : '#f43f5e';

            const customIcon = L.divIcon({
              className: 'custom-div-icon',
              html: `<div style="background-color: ${statusColor}; width: 28px; height: 28px; border-radius: 50%; border: 2.5px solid white; box-shadow: 0 3px 8px rgba(0,0,0,0.5); display: flex; align-items: center; justify-content: center; color: white; font-size: 12px; font-weight: bold;">🏢</div>`,
              iconSize: [28, 28],
              iconAnchor: [14, 14]
            });

            const marker = L.marker([loc.lat, loc.lng], { icon: customIcon });
            marker.bindPopup(`
              <div style="font-family: sans-serif; min-width: 190px;">
                <strong style="color: #4f46e5; font-size: 13px;">${loc.siteId}</strong><br/>
                <div style="font-weight: 600; font-size: 12px; margin-top: 2px;">${loc.name}</div>
                <div style="font-size: 11px; color: #64748b; margin-top: 2px;">📍 ${loc.address}, ${loc.region}</div>
                <div style="font-size: 11px; margin-top: 4px;">Estado: <b>${loc.status}</b></div>
                <button onclick="window.dispatchEvent(new CustomEvent('in-app-route', { detail: '${loc.id}' }))" style="margin-top: 8px; width: 100%; background: #4f46e5; color: white; border: none; padding: 6px; border-radius: 6px; font-size: 11px; font-weight: bold; cursor: pointer;">
                  🧭 Ruta In-App Desde Aquí
                </button>
              </div>
            `);
            markersLayer.addLayer(marker);
          }
        });

        if (targetLocation && targetLocation.lat && targetLocation.lng) {
          map.setView([targetLocation.lat, targetLocation.lng], 16, { animate: false });
        }

        map.invalidateSize();
      }, 100);

      return () => clearTimeout(timer);
    }
  }, [activeTab, filteredLocations, userCoords]);

  // Listen to in-app route event from map popup
  useEffect(() => {
    const handleInAppRoute = (e: any) => {
      const locId = e.detail;
      const found = locations.find(l => l.id === locId);
      if (found) {
        startInAppNavigation(found);
      }
    };
    window.addEventListener('in-app-route', handleInAppRoute as EventListener);
    return () => window.removeEventListener('in-app-route', handleInAppRoute as EventListener);
  }, [locations, userCoords]);

  // AI Insights Generation
  const generateAiInsights = async () => {
    setAiModalOpen(true);
    setAiLoading(true);
    setAiReport(null);

    try {
      const prompt = `Analiza la red de ubicaciones POP FLM (${locations.length} sitios en total). Proporciona un reporte ejecutivo en español con recomendaciones de mantenimiento.`;
      const { GoogleGenAI } = await import('@google/genai');
      const ai = new GoogleGenAI({ apiKey: (import.meta as any).env.VITE_GEMINI_API_KEY || '' });
      const aiResponse = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: prompt
      });
      setAiReport(aiResponse.text || 'Reporte generado.');
    } catch (err) {
      setAiReport(`### Reporte Ejecutivo POP FLM\n\n- **Total Nodos**: ${locations.length} ubicaciones activas.\n- **Sugerencia de Campo**: Monitoreo de enlaces de fibra troncal y respaldos de energía.`);
    } finally {
      setAiLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans selection:bg-indigo-500 selection:text-white pb-20 sm:pb-0">
      {/* Top Header */}
      <header className="bg-slate-900 border-b border-slate-800 sticky top-0 z-30 shadow-xl">
        <div className="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 h-16 sm:h-20 flex items-center justify-between">
          <div className="flex items-center space-x-2.5">
            <div className="w-10 h-10 sm:w-12 sm:h-12 bg-gradient-to-tr from-indigo-600 to-violet-500 rounded-xl flex items-center justify-center shadow-lg shadow-indigo-500/20">
              <Building2 className="w-5 h-5 sm:w-6 sm:h-6 text-white" />
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <h1 className="text-base sm:text-xl font-bold tracking-tight text-white">POP FLM Manager</h1>
                <span className="hidden sm:inline-block px-2.5 py-0.5 text-xs font-semibold bg-indigo-500/20 text-indigo-400 border border-indigo-500/30 rounded-full">
                  v2.0 - 12 Ago
                </span>
              </div>
              <p className="text-[11px] sm:text-xs text-slate-400 truncate max-w-[200px] sm:max-w-md">
                {fileName} {gpsAcquired && <span className="text-emerald-400 ml-1">&bull; GPS Activo</span>}
              </p>
            </div>
          </div>

          <div className="flex items-center space-x-2">
            <button
              onClick={generateAiInsights}
              className="hidden md:flex items-center space-x-2 px-3 py-2 bg-gradient-to-r from-violet-600 to-indigo-600 text-white text-xs sm:text-sm font-medium rounded-xl shadow-lg shadow-indigo-600/30"
            >
              <Sparkles className="w-4 h-4 text-violet-200 animate-pulse" />
              <span>Analizar IA</span>
            </button>

            <button
              onClick={handleExportExcel}
              className="hidden sm:flex items-center space-x-1.5 px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-medium rounded-xl border border-slate-700"
            >
              <Download className="w-4 h-4 text-emerald-400" />
              <span>Exportar</span>
            </button>

            <button
              onClick={() => fileInputRef.current?.click()}
              className="flex items-center space-x-1.5 px-3 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs sm:text-sm font-medium rounded-xl shadow-lg shadow-emerald-600/20"
            >
              <Upload className="w-4 h-4" />
              <span className="hidden sm:inline">Cargar Excel</span>
            </button>
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx, .xls, .csv"
              className="hidden"
            />

            <button
              onClick={handleResetDemo}
              className="p-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl border border-slate-700"
              title="Restaurar Demo"
            >
              <RefreshCw className="w-4 h-4" />
            </button>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 py-4 sm:py-6 flex-1 w-full space-y-4 sm:space-y-6">
        {/* Navigation Tabs */}
        <div className="flex items-center justify-between border-b border-slate-800 pb-3 overflow-x-auto">
          <div className="flex items-center space-x-1.5 bg-slate-900 p-1 rounded-xl border border-slate-800 shrink-0">
            <button
              onClick={() => setActiveTab('table')}
              className={`flex items-center space-x-1.5 px-3 sm:px-4 py-2 rounded-lg text-xs sm:text-sm font-medium transition ${
                activeTab === 'table' ? 'bg-indigo-600 text-white shadow-md' : 'text-slate-400 hover:text-white'
              }`}
            >
              <ListFilter className="w-4 h-4" />
              <span>Ubicaciones ({filteredLocations.length})</span>
            </button>
            <button
              onClick={() => setActiveTab('map')}
              className={`flex items-center space-x-1.5 px-3 sm:px-4 py-2 rounded-lg text-xs sm:text-sm font-medium transition ${
                activeTab === 'map' ? 'bg-indigo-600 text-white shadow-md' : 'text-slate-400 hover:text-white'
              }`}
            >
              <Globe className="w-4 h-4" />
              <span>Mapa & Navegación In-App</span>
            </button>
            <button
              onClick={() => setActiveTab('dashboard')}
              className={`flex items-center space-x-1.5 px-3 sm:px-4 py-2 rounded-lg text-xs sm:text-sm font-medium transition ${
                activeTab === 'dashboard' ? 'bg-indigo-600 text-white shadow-md' : 'text-slate-400 hover:text-white'
              }`}
            >
              <BarChart3 className="w-4 h-4" />
              <span>Dashboard</span>
            </button>
          </div>
        </div>

        {/* TAB 1: TABLE */}
        {activeTab === 'table' && (
          <div className="space-y-4">
            {/* Filters Bar */}
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-3 sm:p-4 shadow-xl grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
              <div className="relative">
                <Search className="absolute left-3.5 top-3 w-4 h-4 text-slate-400" />
                <input
                  type="text"
                  placeholder="Buscar sitio, ID, dirección..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-xl pl-10 pr-4 py-2.5 text-xs sm:text-sm text-slate-100 placeholder-slate-500 focus:outline-none focus:border-indigo-500"
                />
              </div>

              <div>
                <select
                  value={selectedRegion}
                  onChange={(e) => setSelectedRegion(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2.5 text-xs sm:text-sm text-slate-200 focus:outline-none focus:border-indigo-500"
                >
                  <option value="ALL">Todas las Regiones</option>
                  {regions.map(r => <option key={r} value={r}>{r}</option>)}
                </select>
              </div>

              <div>
                <select
                  value={selectedStatus}
                  onChange={(e) => setSelectedStatus(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2.5 text-xs sm:text-sm text-slate-200 focus:outline-none focus:border-indigo-500"
                >
                  <option value="ALL">Todos los Estados</option>
                  {statuses.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>

              <div>
                <select
                  value={selectedTech}
                  onChange={(e) => setSelectedTech(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2.5 text-xs sm:text-sm text-slate-200 focus:outline-none focus:border-indigo-500"
                >
                  <option value="ALL">Todas las Tecnologías</option>
                  {technologies.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </div>
            </div>

            {/* Mobile Cards */}
            <div className="space-y-3 block sm:hidden">
              {paginatedLocations.map(loc => (
                <div key={loc.id} className="bg-slate-900 border border-slate-800 rounded-2xl p-4 shadow-xl space-y-3">
                  <div
                    onClick={() => handleSelectLocationOnMap(loc)}
                    className="cursor-pointer group"
                  >
                    <div className="flex items-start justify-between">
                      <div>
                        <span className="text-xs font-mono font-semibold text-indigo-400 bg-indigo-500/10 px-2 py-0.5 rounded border border-indigo-500/20">
                          {loc.siteId}
                        </span>
                        <h3 className="font-bold text-white text-base mt-1.5 group-hover:text-indigo-300 transition">{loc.name}</h3>
                      </div>
                      <span className={`px-2.5 py-0.5 rounded-full text-[10px] font-bold ${
                        loc.status?.toLowerCase().includes('operativo') ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' :
                        loc.status?.toLowerCase().includes('mantenimiento') ? 'bg-amber-500/10 text-amber-400 border border-amber-500/20' : 'bg-rose-500/10 text-rose-400 border border-rose-500/20'
                      }`}>
                        {loc.status || 'Operativo'}
                      </span>
                    </div>

                    <div className="text-xs text-slate-300 space-y-1 mt-2">
                      <div className="flex items-center gap-1.5 text-indigo-300 underline">
                        <MapPin className="w-3.5 h-3.5 text-indigo-400 shrink-0" />
                        <span className="truncate">{loc.address}, {loc.region} (Ver en Mapa)</span>
                      </div>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-2 pt-2 border-t border-slate-800">
                    <button
                      onClick={() => startInAppNavigation(loc)}
                      className="w-full flex items-center justify-center space-x-1.5 py-2 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-semibold shadow-md"
                    >
                      <Route className="w-3.5 h-3.5" />
                      <span>Ruta In-App</span>
                    </button>
                    <button
                      onClick={() => { setCurrentLocation(loc); setModalMode('view'); }}
                      className="py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-medium flex items-center justify-center gap-1"
                    >
                      <Eye className="w-3.5 h-3.5" /> Detalles
                    </button>
                  </div>
                </div>
              ))}
            </div>

            {/* Desktop Table */}
            <div className="hidden sm:block bg-slate-900 border border-slate-800 rounded-2xl shadow-xl overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-950/80 border-b border-slate-800 text-xs font-semibold text-slate-400 uppercase tracking-wider">
                      <th className="py-4 px-6">ID Sitio</th>
                      <th className="py-4 px-6">Nombre / Ubicación POP (Click para ver logo en mapa)</th>
                      <th className="py-4 px-6">Región / Prov</th>
                      <th className="py-4 px-6">Tecnología</th>
                      <th className="py-4 px-6">Estado</th>
                      <th className="py-4 px-6 text-right">Ruta In-App</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800 text-sm">
                    {paginatedLocations.map((loc) => (
                      <tr key={loc.id} className="hover:bg-slate-800/40 transition">
                        <td className="py-4 px-6 font-mono text-indigo-400 font-medium">
                          {loc.siteId}
                        </td>
                        <td
                          className="py-4 px-6 cursor-pointer group"
                          onClick={() => handleSelectLocationOnMap(loc)}
                        >
                          <div className="font-medium text-slate-100 group-hover:text-indigo-400 transition">{loc.name}</div>
                          <div className="text-xs text-indigo-300 flex items-center gap-1 mt-0.5 underline">
                            <MapPin className="w-3 h-3 text-indigo-400" /> {loc.address}, {loc.region}
                          </div>
                        </td>
                        <td className="py-4 px-6">
                          <div className="text-slate-200">{loc.region}</div>
                        </td>
                        <td className="py-4 px-6">
                          <span className="px-2.5 py-1 text-xs font-medium bg-slate-800 text-slate-300 rounded-lg border border-slate-700">
                            {loc.technology || 'N/A'}
                          </span>
                        </td>
                        <td className="py-4 px-6">
                          <span className={`inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold ${
                            loc.status?.toLowerCase().includes('operativo') ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' :
                            loc.status?.toLowerCase().includes('mantenimiento') ? 'bg-amber-500/10 text-amber-400 border border-amber-500/20' : 'bg-rose-500/10 text-rose-400 border border-rose-500/20'
                          }`}>
                            {loc.status || 'Operativo'}
                          </span>
                        </td>
                        <td className="py-4 px-6 text-right space-x-2">
                          <button
                            onClick={() => startInAppNavigation(loc)}
                            className="px-3 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-medium inline-flex items-center gap-1 shadow"
                          >
                            <Route className="w-3.5 h-3.5" />
                            <span>Ruta In-App</span>
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Pagination */}
            <div className="bg-slate-900 border border-slate-800 px-4 py-3 rounded-2xl flex items-center justify-between">
              <span className="text-xs text-slate-400">
                Página {currentPage} de {totalPages || 1}
              </span>
              <div className="flex space-x-2">
                <button
                  onClick={() => setCurrentPage(p => Math.max(p - 1, 1))}
                  disabled={currentPage === 1}
                  className="px-3 py-1.5 bg-slate-800 disabled:opacity-50 text-slate-200 rounded-lg text-xs font-medium"
                >
                  Anterior
                </button>
                <button
                  onClick={() => setCurrentPage(p => Math.min(p + 1, totalPages))}
                  disabled={currentPage === totalPages}
                  className="px-3 py-1.5 bg-slate-800 disabled:opacity-50 text-slate-200 rounded-lg text-xs font-medium"
                >
                  Siguiente
                </button>
              </div>
            </div>
          </div>
        )}

        {/* TAB 2: IN-APP MAP & NAVIGATION */}
        {activeTab === 'map' && (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
            {/* Map Container (2 cols) */}
            <div className="lg:col-span-2 bg-slate-900 border border-slate-800 rounded-2xl p-4 shadow-xl space-y-3">
              <div className="flex items-center justify-between">
                <div>
                  <h3 className="text-base font-bold text-white flex items-center gap-2">
                    <Globe className="w-5 h-5 text-indigo-400" /> Mapa & Logos de Sites (GPS Activo)
                  </h3>
                  <p className="text-xs text-slate-400">Tus coordenadas GPS y los logos de los sites se muestran en tiempo real.</p>
                </div>
                {navActive && targetLocation && (
                  <button
                    onClick={() => { setNavActive(false); if (routeLayerRef.current && leafletMapInstance.current) routeLayerRef.current.clearLayers(); }}
                    className="px-3 py-1 bg-rose-500/20 text-rose-400 border border-rose-500/30 rounded-xl text-xs font-semibold hover:bg-rose-500/30"
                  >
                    Limpiar Ruta
                  </button>
                )}
              </div>

              <div
                ref={mapRef}
                className="w-full h-[60vh] sm:h-[68vh] rounded-xl border border-slate-800 z-10 shadow-inner bg-slate-950"
              />
            </div>

            {/* In-App Turn-by-Turn Panel (1 col) */}
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 shadow-xl flex flex-col space-y-4">
              <div className="border-b border-slate-800 pb-3">
                <h3 className="text-sm font-bold text-white flex items-center gap-2">
                  <Navigation className="w-4 h-4 text-indigo-400" />
                  Navegación In-App
                </h3>
                <p className="text-xs text-slate-400 mt-0.5">Desde tu GPS hasta el Site seleccionado</p>
              </div>

              {navActive && targetLocation ? (
                <div className="space-y-4 flex-1 flex flex-col">
                  <div className="bg-slate-950 p-3.5 rounded-xl border border-slate-800 space-y-1.5">
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-mono font-semibold text-indigo-400">{targetLocation.siteId}</span>
                      {routeInfo && (
                        <span className="text-xs font-bold text-emerald-400 bg-emerald-500/10 px-2 py-0.5 rounded">
                          {routeInfo.distance} &bull; {routeInfo.duration}
                        </span>
                      )}
                    </div>
                    <div className="font-bold text-white text-sm">{targetLocation.name}</div>
                    <div className="text-xs text-slate-400">📍 {targetLocation.address}, {targetLocation.region}</div>
                  </div>

                  {navLoading ? (
                    <div className="py-12 flex flex-col items-center justify-center space-y-2 text-xs text-slate-400">
                      <div className="w-6 h-6 border-2 border-indigo-500 border-t-transparent rounded-full animate-spin" />
                      <span>Trazando ruta desde tu GPS...</span>
                    </div>
                  ) : routeInfo ? (
                    <div className="space-y-3 flex-1 overflow-y-auto pr-1">
                      <div className="text-xs font-semibold text-slate-300 uppercase tracking-wider">
                        Instrucciones de Ruta:
                      </div>
                      <div className="space-y-2">
                        {routeInfo.steps.map((step, idx) => (
                          <div key={idx} className="bg-slate-950/70 p-3 rounded-xl border border-slate-800/80 flex items-start space-x-3 text-xs text-slate-300">
                            <div className="w-6 h-6 rounded-full bg-indigo-600/20 text-indigo-400 border border-indigo-500/30 flex items-center justify-center shrink-0 font-bold">
                              {idx + 1}
                            </div>
                            <div className="leading-relaxed pt-0.5">{step}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  ) : null}
                </div>
              ) : (
                <div className="flex-1 flex flex-col items-center justify-center text-center p-6 space-y-3 text-slate-500">
                  <Compass className="w-12 h-12 text-slate-700 animate-pulse" />
                  <p className="text-xs">
                    Haz click en cualquier dirección o <strong>Site</strong> de la lista para ver su ubicación en el mapa, o presiona <strong>Ruta In-App</strong> para trazar el recorrido desde tu GPS.
                  </p>
                </div>
              )}
            </div>
          </div>
        )}

        {/* TAB 3: DASHBOARD */}
        {activeTab === 'dashboard' && (
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 shadow-xl">
              <span className="text-sm font-medium text-slate-400">Total Sitios</span>
              <div className="mt-2 text-3xl font-bold text-white">{locations.length}</div>
            </div>
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 shadow-xl">
              <span className="text-sm font-medium text-slate-400">Operativos</span>
              <div className="mt-2 text-3xl font-bold text-emerald-400">
                {locations.filter(l => l.status?.toLowerCase().includes('operativo')).length}
              </div>
            </div>
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 shadow-xl">
              <span className="text-sm font-medium text-slate-400">Mantenimiento</span>
              <div className="mt-2 text-3xl font-bold text-amber-400">
                {locations.filter(l => l.status?.toLowerCase().includes('mantenimiento') || l.status?.toLowerCase().includes('alerta')).length}
              </div>
            </div>
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 shadow-xl">
              <span className="text-sm font-medium text-slate-400">Regiones</span>
              <div className="mt-2 text-3xl font-bold text-violet-400">{regions.length}</div>
            </div>
          </div>
        )}
      </main>

      {/* MOBILE BOTTOM NAVIGATION BAR */}
      <nav className="sm:hidden fixed bottom-0 left-0 right-0 bg-slate-900/95 backdrop-blur-md border-t border-slate-800 z-40 px-6 py-2 flex justify-around items-center">
        <button
          onClick={() => setActiveTab('table')}
          className={`flex flex-col items-center space-y-1 ${activeTab === 'table' ? 'text-indigo-400' : 'text-slate-400'}`}
        >
          <ListFilter className="w-5 h-5" />
          <span className="text-[10px] font-medium">Ubicaciones</span>
        </button>
        <button
          onClick={() => setActiveTab('map')}
          className={`flex flex-col items-center space-y-1 ${activeTab === 'map' ? 'text-indigo-400' : 'text-slate-400'}`}
        >
          <Globe className="w-5 h-5" />
          <span className="text-[10px] font-medium">Mapa In-App</span>
        </button>
        <button
          onClick={() => setActiveTab('dashboard')}
          className={`flex flex-col items-center space-y-1 ${activeTab === 'dashboard' ? 'text-indigo-400' : 'text-slate-400'}`}
        >
          <BarChart3 className="w-5 h-5" />
          <span className="text-[10px] font-medium">Dashboard</span>
        </button>
      </nav>

      {/* VIEW MODAL */}
      {modalMode && currentLocation && (
        <div className="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-lg w-full p-6 shadow-2xl space-y-4">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h3 className="text-base font-bold text-white">Detalles de Ubicación</h3>
              <button onClick={() => setModalMode(null)} className="p-1.5 bg-slate-800 text-slate-400 rounded-lg">
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-3 text-xs sm:text-sm">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-slate-400 mb-1">ID Sitio</label>
                  <div className="font-mono text-indigo-400 font-bold">{currentLocation.siteId}</div>
                </div>
                <div>
                  <label className="block text-slate-400 mb-1">Nombre</label>
                  <div className="font-bold text-white">{currentLocation.name}</div>
                </div>
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-slate-400 mb-1">Región</label>
                  <div>{currentLocation.region}</div>
                </div>
                <div>
                  <label className="block text-slate-400 mb-1">Estado</label>
                  <div>{currentLocation.status}</div>
                </div>
              </div>
              <div>
                <label className="block text-slate-400 mb-1">Dirección</label>
                <div>{currentLocation.address}</div>
              </div>
            </div>

            <div className="flex justify-end pt-3 border-t border-slate-800 space-x-2">
              <button
                onClick={() => {
                  setModalMode(null);
                  handleSelectLocationOnMap(currentLocation);
                }}
                className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl text-xs font-medium"
              >
                Ver en Mapa
              </button>
              <button
                onClick={() => {
                  setModalMode(null);
                  startInAppNavigation(currentLocation);
                }}
                className="px-4 py-2 bg-indigo-600 text-white rounded-xl text-xs font-semibold flex items-center gap-1.5 shadow"
              >
                <Route className="w-4 h-4" /> Ruta In-App
              </button>
            </div>
          </div>
        </div>
      )}

      {/* AI INSIGHTS MODAL */}
      {aiModalOpen && (
        <div className="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-lg w-full p-6 shadow-2xl space-y-4">
            <h3 className="text-base font-bold text-white flex items-center gap-2">
              <Sparkles className="w-5 h-5 text-violet-400 animate-pulse" />
              Análisis Inteligente POP FLM
            </h3>
            <div className="bg-slate-950 p-4 rounded-xl border border-slate-800 text-xs text-slate-300 whitespace-pre-line leading-relaxed max-h-[50vh] overflow-y-auto">
              {aiLoading ? 'Generando análisis de red...' : aiReport}
            </div>
            <div className="flex justify-end">
              <button onClick={() => setAiModalOpen(false)} className="px-4 py-2 bg-indigo-600 text-white rounded-xl text-xs">
                Cerrar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
