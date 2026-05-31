"use client";

import * as Papa from "papaparse";
import { useState, useMemo, useEffect, useRef } from "react";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";

// Safe date parsing helper
function parseDate(dateString: string): string {
  if (!dateString || dateString.trim() === "") {
    return "N/A";
  }
  
  try {
    const mainPart = dateString.split(" GMT")[0];
    const date = new Date(mainPart);
    if (isNaN(date.getTime())) {
      return dateString;
    }
    return date.toISOString().split("T")[0];
  } catch (error) {
    console.warn(`Could not parse date: ${dateString}`, error);
    return dateString;
  }
}

// Inline Custom SVG Icon components to ensure clean aesthetics without external deps
function ShieldLogoIcon() {
  return (
    <svg className="w-10 h-10 text-blue-500 filter drop-shadow-[0_0_8px_rgba(59,130,246,0.6)]" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" />
    </svg>
  );
}

function SearchIcon() {
  return (
    <svg className="w-4 h-4 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
    </svg>
  );
}

function SortIcon({ active, direction }: { active: boolean; direction: "asc" | "desc" }) {
  return (
    <svg className={`w-3.5 h-3.5 transition-colors ${active ? "text-blue-400" : "text-slate-600"}`} fill="currentColor" viewBox="0 0 24 24">
      {direction === "asc" ? (
        <path d="M12 4l-8 8h16l-8-8z" />
      ) : (
        <path d="M12 20l8-8H4l8 8z" />
      )}
    </svg>
  );
}

function CloseIcon() {
  return (
    <svg className="w-5 h-5 text-slate-400 hover:text-white transition-colors" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
    </svg>
  );
}

function MenuIcon() {
  return (
    <svg className="w-6 h-6 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 6h16M4 12h16M4 18h16" />
    </svg>
  );
}

export default function Home() {
  // --- States ---
  const [complaintsData, setComplaintsData] = useState<any[]>([]);
  const [districtsData, setDistrictsData] = useState<any[]>([]);
  const [officesData, setOfficesData] = useState<any[]>([]);
  const [policeStationsData, setPoliceStationsData] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [mobileSidebarOpen, setMobileSidebarOpen] = useState(false);

  // Geographic Select Filters
  const [selectedDistrictId, setSelectedDistrictId] = useState<string | null>(null);
  const [selectedOfficeId, setSelectedOfficeId] = useState<string | null>(null);
  const [selectedPoliceStationId, setSelectedPoliceStationId] = useState<string | null>(null);

  // Advanced Search & Category Filters
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedCategory, setSelectedCategory] = useState("");
  const [statusFilter, setStatusFilter] = useState(""); // "" | "pending" | "resolved" | "investigation"
  
  // Sorting & Interaction
  const [sortConfig, setSortConfig] = useState<{ key: string; direction: "asc" | "desc" } | null>({ key: "date", direction: "desc" });
  const [selectedComplaint, setSelectedComplaint] = useState<any | null>(null);
  const [currentPage, setCurrentPage] = useState(1);

  // Time string to avoid Next.js hydration mismatches
  const [timeString, setTimeString] = useState("");

  const rowsPerPage = 20;

  // File input ref and notification state for data import
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [importStatus, setImportStatus] = useState<string | null>(null);

  // --- Load CSV files on mount ---
  useEffect(() => {
    const loadData = async () => {
      try {
        const [
          complaintsResponse,
          districtsResponse,
          officesResponse,
          policeStationsResponse
        ] = await Promise.all([
          fetch("/data/complaints1.csv"),
          fetch("/data/districts.csv"),
          fetch("/data/offices.csv"),
          fetch("/data/police-stations.csv")
        ]);
        
        const [complaintsText, districtsText, officesText, policeStationsText] = await Promise.all([
          complaintsResponse.text(),
          districtsResponse.text(),
          officesResponse.text(),
          policeStationsResponse.text()
        ]);

        const cleanComplaints = complaintsText.replace(/^\ufeff/, "");
        const cleanDistricts = districtsText.replace(/^\ufeff/, "");
        const cleanOffices = officesText.replace(/^\ufeff/, "");
        const cleanPoliceStations = policeStationsText.replace(/^\ufeff/, "");

        setComplaintsData(Papa.parse(cleanComplaints, { header: true, skipEmptyLines: true }).data);
        setDistrictsData(Papa.parse(cleanDistricts, { header: true, skipEmptyLines: true }).data);
        setOfficesData(Papa.parse(cleanOffices, { header: true, skipEmptyLines: true }).data);
        setPoliceStationsData(Papa.parse(cleanPoliceStations, { header: true, skipEmptyLines: true }).data);
        setLoading(false);
      } catch (error) {
        console.error("Error loading CSV data:", error);
        setLoading(false);
      }
    };
    loadData();
  }, []);

  // Update clock client-side only
  useEffect(() => {
    const timer = setTimeout(() => {
      setTimeString(new Date().toLocaleTimeString());
    }, 0);
    const interval = setInterval(() => {
      setTimeString(new Date().toLocaleTimeString());
    }, 1000);
    return () => {
      clearTimeout(timer);
      clearInterval(interval);
    };
  }, []);

  // --- Mappings & Reference Maps ---
  const districtsMap = useMemo(() => {
    const map = new Map<string, string>();
    districtsData.forEach((d: any) => {
      const key = String(d.id || "").trim();
      if (key) map.set(key, d.name || `District ID: ${key}`);
    });
    return map;
  }, [districtsData]);

  const officesMap = useMemo(() => {
    const map = new Map<string, string>();
    officesData.forEach((o: any) => {
      const key = String(o.id || "").trim();
      if (key) map.set(key, o.name || `Office ID: ${key}`);
    });
    return map;
  }, [officesData]);

  const policeStationsMap = useMemo(() => {
    const map = new Map<string, string>();
    policeStationsData.forEach((ps: any) => {
      const key = String(ps.id || "").trim();
      if (key) map.set(key, ps.name || `Police Station ID: ${key}`);
    });
    officesData.forEach((o: any) => {
      const key = String(o.id || "").trim();
      if (key && !map.has(key)) {
        map.set(key, o.name || `Police Station ID: ${key}`);
      }
    });
    return map;
  }, [policeStationsData, officesData]);

  // --- Hierarchy Relationship Maps ---
  const officesByDistrict = useMemo(() => {
    const map = new Map<string, Set<string>>();
    complaintsData.forEach((complaint: any) => {
      const districtId = String(complaint.transferDistrictCd || complaint.districtMasterId || "").trim();
      const officeId = String(complaint.transferOfficeCd || complaint.officeMasterId || "").trim();
      if (districtId && officeId && officeId !== "0") {
        if (!map.has(districtId)) map.set(districtId, new Set<string>());
        map.get(districtId)?.add(officeId);
      }
    });
    return map;
  }, [complaintsData]);

  const policeStationsByOffice = useMemo(() => {
    const map = new Map<string, Set<string>>();
    complaintsData.forEach((complaint: any) => {
      const officeId = String(complaint.transferOfficeCd || complaint.officeMasterId || "").trim();
      const psId = String(complaint.transferPsCd || complaint.policeStationMasterId || "").trim();
      if (officeId && psId && officeId !== "0" && psId !== "0") {
        if (!map.has(officeId)) map.set(officeId, new Set<string>());
        map.get(officeId)?.add(psId);
      }
    });
    return map;
  }, [complaintsData]);

  const policeStationsFromComplaints = useMemo(() => {
    const map = new Map<string, Set<string>>();
    complaintsData.forEach((complaint: any) => {
      const districtId = String(complaint.transferDistrictCd || complaint.districtMasterId || "").trim();
      const psId = String(complaint.transferPsCd || complaint.policeStationMasterId || "").trim();
      if (districtId && psId && psId !== "0") {
        if (!map.has(districtId)) map.set(districtId, new Set<string>());
        map.get(districtId)?.add(psId);
      }
    });
    return map;
  }, [complaintsData]);

  // --- List Population for Filters ---
  const uniqueDistricts = useMemo(() => {
    const districtIds = new Set<string>();
    complaintsData.forEach((c: any) => {
      const id = String(c.transferDistrictCd || c.districtMasterId || "").trim();
      if (id) districtIds.add(id);
    });
    return Array.from(districtIds)
      .map(id => ({ 
        id, 
        name: districtsMap.get(id) || `District ID: ${id}`
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [complaintsData, districtsMap]);

  const filteredOffices = useMemo(() => {
    if (!selectedDistrictId) return [];
    const officeIds = officesByDistrict.get(selectedDistrictId) || new Set();
    return Array.from(officeIds)
      .map(id => ({ 
        id, 
        name: officesMap.get(id) || `Office ID: ${id}`
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [selectedDistrictId, officesByDistrict, officesMap]);

  const filteredPoliceStations = useMemo(() => {
    if (selectedOfficeId) {
      const psIds = policeStationsByOffice.get(selectedOfficeId) || new Set();
      return Array.from(psIds)
        .map(id => ({ 
          id, 
          name: policeStationsMap.get(id) || `Police Station ID: ${id}`
        }))
        .sort((a, b) => a.name.localeCompare(b.name));
    } else if (selectedDistrictId) {
      const psIds = policeStationsFromComplaints.get(selectedDistrictId) || new Set();
      return Array.from(psIds)
        .map(id => ({ 
          id, 
          name: policeStationsMap.get(id) || `Police Station ID: ${id}`
        }))
        .sort((a, b) => a.name.localeCompare(b.name));
    }
    return [];
  }, [selectedDistrictId, selectedOfficeId, policeStationsByOffice, policeStationsFromComplaints, policeStationsMap]);

  // Dynamic Incident Categories populated from data
  const uniqueCategories = useMemo(() => {
    const categoriesSet = new Set<string>();
    complaintsData.forEach((c: any) => {
      const cat = String(c.incidentType || "").trim();
      if (cat) categoriesSet.add(cat);
    });
    return Array.from(categoriesSet).sort();
  }, [complaintsData]);

  // --- Filtering Logic ---
  const filteredComplaints = useMemo(() => {
    return complaintsData.filter((complaint: any) => {
      const districtId = String(complaint.transferDistrictCd || complaint.districtMasterId || "").trim();
      const officeId = String(complaint.transferOfficeCd || complaint.officeMasterId || "").trim();
      const policeStationId = String(complaint.transferPsCd || complaint.policeStationMasterId || "").trim();
      
      const effectiveOfficeId = (officeId && officeId !== "0") ? officeId : "";
      
      const matchesDistrict = !selectedDistrictId || districtId === selectedDistrictId;
      const matchesOffice = !selectedOfficeId || effectiveOfficeId === selectedOfficeId;
      const matchesPoliceStation = !selectedPoliceStationId || policeStationId === selectedPoliceStationId;

      // Text Search Query (Complainant Name, ID, or Phone, Description)
      let matchesSearch = true;
      if (searchQuery.trim() !== "") {
        const q = searchQuery.toLowerCase().trim();
        const complId = String(complaint.complRegNum || "").toLowerCase();
        const firstName = String(complaint.firstName || "").toLowerCase();
        const lastName = String(complaint.lastName || "").toLowerCase();
        const fullName = `${firstName} ${lastName}`;
        const desc = String(complaint.complDesc || "").toLowerCase();
        const mobile = String(complaint.mobile || "").toLowerCase();
        
        matchesSearch = complId.includes(q) || 
                        firstName.includes(q) || 
                        lastName.includes(q) || 
                        fullName.includes(q) || 
                        desc.includes(q) || 
                        mobile.includes(q);
      }

      // Dropdown Category Filter
      let matchesCategory = true;
      if (selectedCategory !== "") {
        matchesCategory = String(complaint.incidentType || "").toLowerCase() === selectedCategory.toLowerCase();
      }

      // Status Pill Group Filter
      let matchesStatus = true;
      if (statusFilter !== "") {
        const status = String(complaint.statusGroup || complaint.statusOfComplaint || complaint.statusRaw || "").toLowerCase();
        if (statusFilter === "pending") {
          matchesStatus = status === "pending";
        } else if (statusFilter === "resolved") {
          matchesStatus = status === "disposed" || status === "resolved" || status === "closed";
        } else if (statusFilter === "investigation") {
          matchesStatus = status.includes("investigation");
        }
      }

      return matchesDistrict && matchesOffice && matchesPoliceStation && matchesSearch && matchesCategory && matchesStatus;
    });
  }, [selectedDistrictId, selectedOfficeId, selectedPoliceStationId, searchQuery, selectedCategory, statusFilter, complaintsData]);

  // --- Sorting Logic ---
  const sortedComplaints = useMemo(() => {
    if (!sortConfig) return filteredComplaints;
    
    return [...filteredComplaints].sort((a: any, b: any) => {
      let aVal: any = "";
      let bVal: any = "";
      
      const key = sortConfig.key;
      if (key === "id") {
        const aNum = Number(a.complRegNum || 0);
        const bNum = Number(b.complRegNum || 0);
        aVal = isNaN(aNum) ? 0 : aNum;
        bVal = isNaN(bNum) ? 0 : bNum;
      } else if (key === "date") {
        aVal = a.complRegDt ? new Date(a.complRegDt).getTime() : 0;
        bVal = b.complRegDt ? new Date(b.complRegDt).getTime() : 0;
      } else if (key === "complainant") {
        aVal = `${a.firstName || ""} ${a.lastName || ""}`.trim().toLowerCase();
        bVal = `${b.firstName || ""} ${b.lastName || ""}`.trim().toLowerCase();
      } else if (key === "status") {
        aVal = String(a.statusGroup || a.statusOfComplaint || a.statusRaw || "").toLowerCase();
        bVal = String(b.statusGroup || b.statusOfComplaint || b.statusRaw || "").toLowerCase();
      } else if (key === "district") {
        const dIdA = String(a.transferDistrictCd || a.districtMasterId || "").trim();
        const dIdB = String(b.transferDistrictCd || b.districtMasterId || "").trim();
        aVal = districtsMap.get(dIdA) || "";
        bVal = districtsMap.get(dIdB) || "";
      } else {
        aVal = String(a[key] || "").toLowerCase();
        bVal = String(b[key] || "").toLowerCase();
      }
      
      if (aVal < bVal) return sortConfig.direction === "asc" ? -1 : 1;
      if (aVal > bVal) return sortConfig.direction === "asc" ? 1 : -1;
      return 0;
    });
  }, [filteredComplaints, sortConfig, districtsMap]);

  // --- Analytics Derivations ---
  const totalCount = filteredComplaints.length;

  const pendingCount = useMemo(() => {
    return filteredComplaints.filter((c: any) => {
      const s = String(c.statusGroup || c.statusOfComplaint || c.statusRaw || "").toLowerCase();
      return s === "pending";
    }).length;
  }, [filteredComplaints]);

  const resolvedCount = useMemo(() => {
    return filteredComplaints.filter((c: any) => {
      const s = String(c.statusGroup || c.statusOfComplaint || c.statusRaw || "").toLowerCase();
      return s === "disposed" || s === "resolved" || s === "closed";
    }).length;
  }, [filteredComplaints]);

  const investigationCount = useMemo(() => {
    return filteredComplaints.filter((c: any) => {
      const s = String(c.statusGroup || c.statusOfComplaint || c.statusRaw || "").toLowerCase();
      return s.includes("investigation");
    }).length;
  }, [filteredComplaints]);

  const resolutionRate = useMemo(() => {
    if (totalCount === 0) return 0;
    return Math.round((resolvedCount / totalCount) * 100);
  }, [totalCount, resolvedCount]);

  // Dynamic Horizontal Bars Chart Data: Top 5 Categories
  const categoryChartData = useMemo(() => {
    const counts: Record<string, number> = {};
    filteredComplaints.forEach((c: any) => {
      const cat = String(c.incidentType || "Unknown/Other").trim();
      counts[cat] = (counts[cat] || 0) + 1;
    });
    return Object.entries(counts)
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);
  }, [filteredComplaints]);

  // Submission Mode Stats
  const submissionModeData = useMemo(() => {
    const counts: Record<string, number> = {};
    filteredComplaints.forEach((c: any) => {
      const mode = String(c.receptionMode || "Direct/Self").trim();
      counts[mode] = (counts[mode] || 0) + 1;
    });
    return Object.entries(counts)
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 4);
  }, [filteredComplaints]);

  // --- Pagination Logic ---
  const totalPages = Math.ceil(totalCount / rowsPerPage);
  
  const tableData = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return sortedComplaints.slice(start, start + rowsPerPage).map((c: any) => {
      const districtId = String(c.transferDistrictCd || c.districtMasterId || "").trim();
      const officeId = String(c.transferOfficeCd || c.officeMasterId || "").trim();
      const policeStationId = String(c.transferPsCd || c.policeStationMasterId || "").trim();
      
      let complaintId = "N/A";
      if (c.complRegNum !== undefined && c.complRegNum !== null && c.complRegNum !== "") {
        const raw = String(c.complRegNum).trim();
        try {
          if (raw.includes('E') || raw.includes('e')) {
            const num = Number(raw);
            if (!isNaN(num)) {
              complaintId = BigInt(Math.round(num)).toString();
            } else {
              complaintId = raw;
            }
          } else {
            complaintId = raw;
          }
        } catch {
          complaintId = raw;
        }
      }
      
      const districtName = districtId ? (districtsMap.get(districtId) || `District ID: ${districtId}`) : "N/A";
      const officeName = officeId && officeId !== "0" ? (officesMap.get(officeId) || `Office ID: ${officeId}`) : "N/A";
      const policeStationName = policeStationId && policeStationId !== "0" ? (policeStationsMap.get(policeStationId) || `PS ID: ${policeStationId}`) : "N/A";
      
      const status = String(c.statusGroup || c.statusOfComplaint || c.statusRaw || "Unknown").trim();
      const category = String(c.incidentType || "Unknown").trim();
      const complainant = `${c.firstName || ""} ${c.lastName || ""}`.trim() || "Unknown";
                          
      return {
        id: complaintId,
        district: districtName,
        office: officeName,
        policeStation: policeStationName,
        status,
        date: parseDate(c.complRegDt),
        category,
        complainant,
        raw: c // reference for slide drawer
      };
    });
  }, [sortedComplaints, districtsMap, officesMap, policeStationsMap, currentPage]);

  // --- Handlers ---
  const handleDistrictChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const value = e.target.value;
    setSelectedDistrictId(value === "" ? null : value);
    setSelectedOfficeId(null);
    setSelectedPoliceStationId(null);
    setCurrentPage(1);
  };

  const handleOfficeChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const value = e.target.value;
    setSelectedOfficeId(value === "" ? null : value);
    setSelectedPoliceStationId(null);
    setCurrentPage(1);
  };

  const handlePoliceStationChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const value = e.target.value;
    setSelectedPoliceStationId(value === "" ? null : value);
    setCurrentPage(1);
  };

  const toggleSort = (key: string) => {
    if (sortConfig && sortConfig.key === key) {
      setSortConfig({
        key,
        direction: sortConfig.direction === "asc" ? "desc" : "asc"
      });
    } else {
      setSortConfig({ key, direction: "desc" });
    }
    setCurrentPage(1);
  };

  // --- Import / Export Handlers ---
  const downloadSampleTemplate = () => {
    const headers = [
      "complRegNum", "complRegDt", "districtName", "districtMasterId", "policeStationMasterId",
      "officeMasterId", "complDesc", "complSrno", "firstName", "lastName", "mobile", "gender",
      "age", "addressLine1", "addressLine2", "addressLine3", "village", "tehsil", "addressDistrict",
      "addressPs", "receptionMode", "incidentType", "incidentPlc", "incidentFromDt", "incidentToDt",
      "submitPsCd", "submitOfficeCd", "email", "statusRaw", "statusGroup", "statusOfComplaint",
      "disposalDate", "classOfIncident", "complaintSource", "typeOfComplaint", "crimeCategory",
      "complainantType", "complaintPurpose", "ioDetails", "respondentCategories", "transferDistrictCd",
      "transferOfficeCd", "transferPsCd"
    ];
    
    const sampleRow1 = [
      "130000000001", "2026-04-01T00:00:00.000Z", "Karnal", "13231", "13231011",
      "130204", "The complainant has reported a property boundary dispute with neighbor.", "1101", "Rajesh", "Kumar", "9872100001", "Male",
      "35", "123", "Sector 14", "Near Main Market", "Karnal", "Karnal", "KARNAL",
      "Karnal", "In-Person/By Hand", "Property/Land Dispute", "Karnal", "2026-03-31T00:00:00.000Z", "2026-04-01T00:00:00.000Z",
      "13231011", "130204", "rajesh.kumar@email.com", "pending", "pending", "pending",
      "", "Property/Land Dispute", "Citizen/General Public", "Fresh complaint", "Property/Land Dispute",
      "Private person", "Enquiry", "Suresh Kumar/SI/9876543211", "Against Private Person", "13231", "130204", "13231011"
    ];

    const sampleRow2 = [
      "130000000002", "2026-04-02T00:00:00.000Z", "Gurugram", "13227", "132270403",
      "130204", "Online payment fraud of Rs. 50,000 from credit card.", "1102", "Anita", "Sharma", "9998887776", "Female",
      "29", "Flat 402", "Cyber City", "Phase 3", "Gurugram", "Gurugram", "GURUGRAM",
      "Cyber City", "Online", "Other Economic Offence", "Gurugram", "2026-04-01T00:00:00.000Z", "2026-04-02T00:00:00.000Z",
      "132270403", "130204", "anita.sharma@email.com", "disposed", "resolved", "Disposed- Resolved",
      "2026-04-05T00:00:00.000Z", "Other Economic Offence", "Citizen/General Public", "Fresh complaint", "Other Economic Offence",
      "Private person", "Enquiry", "Ramesh Chander/Inspector/9998887775", "Against Private Person", "13227", "130204", "132270403"
    ];

    const csvContent = [
      headers.join(","),
      sampleRow1.map(val => `"${val.replace(/"/g, '""')}"`).join(","),
      sampleRow2.map(val => `"${val.replace(/"/g, '""')}"`).join(",")
    ].join("\n");

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", "complaints_import_template.csv");
    link.style.visibility = "hidden";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const processImportedData = (data: any[]) => {
    if (!data || data.length === 0) {
      alert("No data found in the uploaded file.");
      return;
    }

    const sampleRow = data[0];
    const essentialKeys = ["complRegNum", "firstName"];
    const hasAnyEssential = essentialKeys.some(key => key in sampleRow);
    
    if (!hasAnyEssential) {
      alert("Validation failed: The file does not match the template layout. Please download the sample template first.");
      return;
    }

    const cleanedRows = data.map((row: any, idx: number) => {
      return {
        ...row,
        complRegNum: row.complRegNum ? String(row.complRegNum).trim() : `IMP-${Date.now()}-${idx}`,
        complRegDt: row.complRegDt ? String(row.complRegDt).trim() : new Date().toISOString(),
        statusGroup: row.statusGroup ? String(row.statusGroup).trim().toLowerCase() : "pending",
        firstName: row.firstName ? String(row.firstName).trim() : "Unknown",
        lastName: row.lastName ? String(row.lastName).trim() : "",
      };
    });

    setComplaintsData(prev => [...cleanedRows, ...prev]);
    setImportStatus(`Successfully imported ${cleanedRows.length} cases.`);
    setCurrentPage(1);
    
    setTimeout(() => {
      setImportStatus(null);
    }, 5000);
  };

  const handleImportFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const fileExtension = file.name.split('.').pop()?.toLowerCase();
    
    if (fileExtension === "csv") {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          processImportedData(results.data);
        },
        error: (error) => {
          alert(`Error parsing CSV: ${error.message}`);
        }
      });
    } else if (fileExtension === "xlsx" || fileExtension === "xls") {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const bstr = evt.target?.result;
          const workbook = XLSX.read(bstr, { type: "binary" });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const data = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
          processImportedData(data);
        } catch (error: any) {
          alert(`Error parsing Excel file: ${error?.message || error}`);
        }
      };
      reader.readAsBinaryString(file);
    } else {
      alert("Unsupported file format. Please upload a .csv or .xlsx/.xls file.");
    }
    
    e.target.value = "";
  };

  const exportPDF = () => {
    // Landscape A4 size (297 x 210 mm)
    const doc = new jsPDF("l", "mm", "a4");

    const currentDateTime = new Date().toLocaleString("en-US", {
      dateStyle: "medium",
      timeStyle: "short"
    });

    // Limit export to 1000 rows max
    const maxRows = 1000;
    const dataToExport = sortedComplaints.slice(0, maxRows);

    const headers = [["Complaint ID", "District", "Office", "Police Station", "Status", "Date", "Category"]];
    
    const body = dataToExport.map((c: any) => {
      let complaintId = "N/A";
      if (c.complRegNum !== undefined && c.complRegNum !== null && c.complRegNum !== "") {
        const raw = String(c.complRegNum).trim();
        try {
          if (raw.includes('E') || raw.includes('e')) {
            const num = Number(raw);
            complaintId = !isNaN(num) ? BigInt(Math.round(num)).toString() : raw;
          } else {
            complaintId = raw;
          }
        } catch {
          complaintId = raw;
        }
      }

      const districtId = String(c.transferDistrictCd || c.districtMasterId || "").trim();
      const officeId = String(c.transferOfficeCd || c.officeMasterId || "").trim();
      const policeStationId = String(c.transferPsCd || c.policeStationMasterId || "").trim();

      const districtName = districtId ? (districtsMap.get(districtId) || `ID: ${districtId}`) : "N/A";
      const officeName = officeId && officeId !== "0" ? (officesMap.get(officeId) || `ID: ${officeId}`) : "N/A";
      const policeStationName = policeStationId && policeStationId !== "0" ? (policeStationsMap.get(policeStationId) || `ID: ${policeStationId}`) : "N/A";
      
      const status = String(c.statusGroup || c.statusOfComplaint || c.statusRaw || "Unknown").toUpperCase();
      const category = String(c.incidentType || "Unknown").toUpperCase();
      
      return [
        complaintId,
        districtName,
        officeName,
        policeStationName,
        status,
        parseDate(c.complRegDt),
        category
      ];
    });
    
    // Header
    doc.setFont("helvetica", "bold");
    doc.setFontSize(20);
    doc.setTextColor(11, 23, 44);
    doc.text("Police Headquarters Complaint Report", 14, 18);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.setTextColor(80, 80, 80);
    doc.text(`Generated On: ${currentDateTime}`, 14, 25);
    doc.text(`Total Filtered Complaints: ${filteredComplaints.length} (Showing first ${dataToExport.length} rows)`, 14, 30);

    const filtersApplied = [];
    if (selectedDistrictId) filtersApplied.push(`District: ${districtsMap.get(selectedDistrictId) || selectedDistrictId}`);
    if (selectedOfficeId) filtersApplied.push(`Office: ${officesMap.get(selectedOfficeId) || selectedOfficeId}`);
    if (selectedPoliceStationId) filtersApplied.push(`PS: ${policeStationsMap.get(selectedPoliceStationId) || selectedPoliceStationId}`);
    if (selectedCategory) filtersApplied.push(`Category: ${selectedCategory}`);
    if (statusFilter) filtersApplied.push(`Status: ${statusFilter.toUpperCase()}`);
    if (searchQuery) filtersApplied.push(`Search: "${searchQuery}"`);
    
    const filterText = filtersApplied.length > 0 ? filtersApplied.join(" | ") : "None (All complaints)";
    doc.text(`Active Filters: ${filterText}`, 14, 35);

    doc.setDrawColor(200, 200, 200);
    doc.setLineWidth(0.5);
    doc.line(14, 38, 283, 38);

    autoTable(doc, {
      head: headers,
      body: body,
      startY: 42,
      margin: { left: 14, right: 14 },
      theme: "striped",
      headStyles: {
        fillColor: [11, 23, 44],
        textColor: [255, 255, 255],
        fontSize: 8,
        fontStyle: "bold",
        halign: "left"
      },
      bodyStyles: {
        fontSize: 8,
        textColor: [50, 50, 50],
        cellPadding: 3
      },
      alternateRowStyles: {
        fillColor: [245, 247, 250]
      },
      columnStyles: {
        0: { cellWidth: 32 },
        1: { cellWidth: 32 },
        2: { cellWidth: 40 },
        3: { cellWidth: 40 },
        4: { cellWidth: 28 },
        5: { cellWidth: 25 },
        6: { cellWidth: "auto" }
      },
      didDrawPage: () => {
        const str = `Page ${doc.getNumberOfPages()}`;
        doc.setFontSize(8);
        doc.setTextColor(120, 120, 120);
        doc.text(str, 283 - doc.getTextWidth(str), 205);
        doc.text("CONFIDENTIAL - HARYANA POLICE STATEWIDE COMPLAINTS SYSTEM", 14, 205);
      }
    });

    const timestamp = new Date().toISOString().replace(/[-:.]/g, "").slice(0, 15);
    doc.save(`phq_complaint_report_${timestamp}.pdf`);
  };

  const renderSidebarContent = () => {
    return (
      <>
        <div className="p-6 space-y-8 flex-1">
          
          {/* Main Logo Crest */}
          <div className="flex items-center space-x-3.5 pb-4 border-b border-police-700/30">
            <ShieldLogoIcon />
            <div>
              <h2 className="text-md font-extrabold tracking-wider bg-clip-text text-transparent bg-gradient-to-r from-blue-400 to-indigo-200">HARYANA POLICE</h2>
              <p className="text-[10px] text-blue-400 tracking-widest uppercase">HQ Control Room</p>
            </div>
          </div>

          {/* Geographical Filters Section */}
          <div className="space-y-4">
            <h4 className="text-xs font-semibold tracking-wider text-slate-400 uppercase flex items-center">
              <span className="w-1.5 h-1.5 bg-blue-500 rounded-full mr-2"></span>
              Geographic Sector
            </h4>
            
            <div className="space-y-3.5">
              <div className="space-y-1">
                <label className="text-[11px] text-slate-400">District Master</label>
                <select
                  value={selectedDistrictId ?? ""}
                  onChange={handleDistrictChange}
                  className="w-full bg-[#05070f] border border-police-700/50 rounded-lg px-3.5 py-2 text-sm text-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-blue-500/50 transition-all cursor-pointer"
                >
                  <option value="">[ All Districts ]</option>
                  {uniqueDistricts.map(district => (
                    <option key={district.id} value={district.id}>{district.name}</option>
                  ))}
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-[11px] text-slate-400">HQ/Office Master</label>
                <select
                  value={selectedOfficeId ?? ""}
                  onChange={handleOfficeChange}
                  disabled={!selectedDistrictId}
                  className="w-full bg-[#05070f] border border-police-700/50 rounded-lg px-3.5 py-2 text-sm text-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-blue-500/50 disabled:opacity-30 disabled:cursor-not-allowed transition-all cursor-pointer"
                >
                  <option value="">[ All Offices ]</option>
                  {filteredOffices.map(office => (
                    <option key={office.id} value={office.id}>{office.name}</option>
                  ))}
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-[11px] text-slate-400">Police Station</label>
                <select
                  value={selectedPoliceStationId ?? ""}
                  onChange={handlePoliceStationChange}
                  disabled={!selectedDistrictId}
                  className="w-full bg-[#05070f] border border-police-700/50 rounded-lg px-3.5 py-2 text-sm text-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-blue-500/50 disabled:opacity-30 disabled:cursor-not-allowed transition-all cursor-pointer"
                >
                  <option value="">[ All Police Stations ]</option>
                  {filteredPoliceStations.map(ps => (
                    <option key={ps.id} value={ps.id}>{ps.name}</option>
                  ))}
                </select>
              </div>
            </div>
          </div>

          {/* Quick Filters Status pills */}
          <div className="space-y-3.5">
            <h4 className="text-xs font-semibold tracking-wider text-slate-400 uppercase flex items-center">
              <span className="w-1.5 h-1.5 bg-indigo-500 rounded-full mr-2"></span>
              Incident Status
            </h4>
            <div className="flex flex-col gap-2">
              <button
                onClick={() => { setStatusFilter(""); setCurrentPage(1); setMobileSidebarOpen(false); }}
                className={`w-full text-left px-3 py-2 rounded-lg text-xs transition-all flex justify-between items-center ${
                  statusFilter === "" 
                    ? "bg-blue-600/20 text-blue-400 border border-blue-500/40 shadow-[inset_0_0_8px_rgba(59,130,246,0.15)] font-semibold" 
                    : "bg-[#05070f]/50 hover:bg-[#0c1228]/50 text-slate-400 border border-transparent"
                }`}
              >
                <span>ALL COMPLAINTS</span>
                <span className="bg-slate-800 text-slate-300 px-1.5 py-0.5 rounded text-[10px]">{complaintsData.length}</span>
              </button>
              
              <button
                onClick={() => { setStatusFilter("pending"); setCurrentPage(1); setMobileSidebarOpen(false); }}
                className={`w-full text-left px-3 py-2 rounded-lg text-xs transition-all flex justify-between items-center ${
                  statusFilter === "pending"
                    ? "bg-amber-600/20 text-amber-400 border border-amber-500/40 shadow-[inset_0_0_8px_rgba(245,158,11,0.15)] font-semibold" 
                    : "bg-[#05070f]/50 hover:bg-[#0c1228]/50 text-slate-400 border border-transparent"
                }`}
              >
                <span>PENDING TRIBUNAL</span>
                <span className="bg-amber-950 text-amber-400 border border-amber-900/50 px-1.5 py-0.5 rounded text-[10px]">Pending</span>
              </button>

              <button
                onClick={() => { setStatusFilter("investigation"); setCurrentPage(1); setMobileSidebarOpen(false); }}
                className={`w-full text-left px-3 py-2 rounded-lg text-xs transition-all flex justify-between items-center ${
                  statusFilter === "investigation"
                    ? "bg-indigo-600/20 text-indigo-400 border border-indigo-500/40 shadow-[inset_0_0_8px_rgba(99,102,241,0.15)] font-semibold" 
                    : "bg-[#05070f]/50 hover:bg-[#0c1228]/50 text-slate-400 border border-transparent"
                }`}
              >
                <span>UNDER INQUIRY</span>
                <span className="bg-indigo-950 text-indigo-400 border border-indigo-900/50 px-1.5 py-0.5 rounded text-[10px]">Inquiry</span>
              </button>

              <button
                onClick={() => { setStatusFilter("resolved"); setCurrentPage(1); setMobileSidebarOpen(false); }}
                className={`w-full text-left px-3 py-2 rounded-lg text-xs transition-all flex justify-between items-center ${
                  statusFilter === "resolved"
                    ? "bg-emerald-600/20 text-emerald-400 border border-emerald-500/40 shadow-[inset_0_0_8px_rgba(16,185,129,0.15)] font-semibold" 
                    : "bg-[#05070f]/50 hover:bg-[#0c1228]/50 text-slate-400 border border-transparent"
                }`}
              >
                <span>DISPOSED / CLOSED</span>
                <span className="bg-emerald-950 text-emerald-400 border border-emerald-900/50 px-1.5 py-0.5 rounded text-[10px]">Disposed</span>
              </button>
            </div>
          </div>
        </div>

        {/* System info footer */}
        <div className="p-4 border-t border-police-700/30 bg-[#05070f]/80 text-[10px] text-slate-400 space-y-1">
          <p>CONSOLE_VER: 4.8.2-SECURE</p>
          <p>DB_TARGET: Neon.PostgreSQL</p>
          <p>UPDATED: {timeString || "LOADING..."}</p>
        </div>
      </>
    );
  };

  // Radial Ring Circle Parameters
  const ringRadius = 42;
  const ringCircumference = 2 * Math.PI * ringRadius;
  const ringOffset = ringCircumference - (resolutionRate / 100) * ringCircumference;

  // --- Loading screen ---
  if (loading) {
    return (
      <div className="min-h-screen bg-[#04060f] flex items-center justify-center">
        <div className="text-center space-y-4">
          <div className="relative inline-flex">
            <div className="w-16 h-16 rounded-full border-4 border-blue-500/20 border-t-blue-500 animate-spin"></div>
            <div className="absolute inset-0 flex items-center justify-center">
              <div className="w-8 h-8 bg-blue-500/10 rounded-full animate-ping"></div>
            </div>
          </div>
          <h3 className="text-lg font-semibold text-slate-300 tracking-wider">HARYANA POLICE</h3>
          <p className="text-xs text-slate-500 tracking-widest animate-pulse">BOOTING DECISION ENGINE...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-[#060814] via-[#090d22] to-[#04060f] text-slate-100 font-sans flex flex-col lg:flex-row overflow-hidden relative">
      
      {/* Desktop Sidebar Control Panel */}
      <aside className="w-80 glass-panel border-r border-police-700/30 flex-col shrink-0 z-20 h-screen overflow-y-auto hidden lg:flex">
        {renderSidebarContent()}
      </aside>

      {/* Mobile Drawer Sidebar Overlay */}
      {mobileSidebarOpen && (
        <div className="fixed inset-0 z-50 flex lg:hidden">
          {/* Backdrop */}
          <div 
            onClick={() => setMobileSidebarOpen(false)}
            className="absolute inset-0 bg-[#04060f]/60 backdrop-blur-sm animate-fade-in"
          />
          {/* Drawer content sliding in from left */}
          <aside className="w-80 bg-[#070b1b] border-r border-police-700/35 h-full relative z-10 flex flex-col shadow-[0_0_50px_rgba(0,0,0,0.8)] animate-slide-in-left">
            <div className="flex justify-end p-4 border-b border-police-700/30">
              <button 
                onClick={() => setMobileSidebarOpen(false)}
                className="p-2 bg-police-950 rounded-lg hover:bg-police-800 border border-police-700/40 cursor-pointer transition-all"
              >
                <CloseIcon />
              </button>
            </div>
            <div className="flex-1 overflow-y-auto flex flex-col">
              {renderSidebarContent()}
            </div>
          </aside>
        </div>
      )}

      {/* Main Command Center Body */}
      <main className="flex-1 flex flex-col h-screen overflow-y-auto z-10 relative">
        
        {/* Top Header Navigation Bar */}
        <header className="glass-panel border-b border-police-700/20 shrink-0 px-4 py-4 md:px-8 md:py-5 flex flex-col md:flex-row md:items-center justify-between gap-4 sticky top-0 bg-gradient-to-r from-[#060814]/90 to-[#0a0f26]/90 z-30">
          <div className="flex items-center gap-3">
            <button
              onClick={() => setMobileSidebarOpen(true)}
              className="lg:hidden p-2 rounded-lg bg-[#05070f] border border-police-700/40 text-slate-400 hover:text-white transition-all cursor-pointer"
              title="Open Controls Menu"
            >
              <MenuIcon />
            </button>
            <div>
              <h1 className="text-base md:text-xl font-black tracking-wide text-white uppercase flex items-center gap-2">
                Complaint Operations Console
                <span className="bg-red-950/80 text-red-500 border border-red-900/50 px-1.5 py-0.5 md:px-2 md:py-0.5 rounded text-[8px] md:text-[10px] tracking-widest font-bold animate-pulse">LIVE</span>
              </h1>
              <p className="text-[10px] md:text-xs text-slate-400">HARYANA POLICE STATEWIDE COMPLAINTS & INVESTIGATIONS DATABASE</p>
            </div>
          </div>
          
          <div className="flex flex-wrap items-center gap-2 md:gap-3 w-full md:w-auto">
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleImportFile}
              accept=".csv,.xlsx,.xls"
              className="hidden"
            />
            <button
              onClick={() => fileInputRef.current?.click()}
              className="px-3 py-1.5 text-xs font-bold bg-blue-600 hover:bg-blue-700 text-white rounded-lg transition-all shadow-[0_0_10px_rgba(59,130,246,0.3)] cursor-pointer flex items-center gap-1.5 border border-blue-500"
            >
              <span>📥</span> Import Data
            </button>
            <button
              onClick={downloadSampleTemplate}
              className="px-3 py-1.5 text-xs font-semibold bg-police-950 hover:bg-police-900 border border-police-700/40 text-slate-300 hover:text-white rounded-lg transition-all cursor-pointer flex items-center gap-1.5"
              title="Download CSV sample template for bulk import"
            >
              <span>📄</span> Template
            </button>
 
            <div className="hidden sm:flex items-center space-x-2 bg-police-950 px-3 py-1.5 rounded-lg border border-police-700/40">
              <span className="w-2.5 h-2.5 bg-emerald-500 rounded-full animate-ping"></span>
              <span className="w-2.5 h-2.5 bg-emerald-500 rounded-full absolute"></span>
              <span className="text-[11px] text-slate-300 tracking-wider font-semibold">CON_SECURE: CONNECTED</span>
            </div>
          </div>
        </header>

        <div className="px-4 py-6 md:px-8 md:py-8 space-y-6 md:space-y-8 flex-1">
          
          {/* Analytical summary cards */}
          <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 md:gap-5">
            
            {/* Total complaints */}
            <div className="glass-panel glass-panel-hover rounded-xl p-5 relative overflow-hidden flex flex-col justify-between h-32 shadow-[0_4px_20px_rgba(0,0,0,0.15)]">
              <div className="flex items-center justify-between">
                <span className="text-xs text-slate-400 uppercase tracking-widest">Active Cases</span>
                <span className="w-7 h-7 bg-blue-500/10 text-blue-400 rounded-lg flex items-center justify-center border border-blue-500/30">⚡</span>
              </div>
              <div className="space-y-1">
                <h3 className="text-3xl font-black text-white leading-none tracking-tight">{totalCount.toLocaleString()}</h3>
                <p className="text-[10px] text-slate-400">Filtered size of case registry</p>
              </div>
              <div className="absolute right-0 bottom-0 translate-x-2 translate-y-3 opacity-[0.03] select-none pointer-events-none">
                <ShieldLogoIcon />
              </div>
            </div>

            {/* Pending complaints */}
            <div className="glass-panel glass-panel-hover rounded-xl p-5 relative overflow-hidden flex flex-col justify-between h-32 shadow-[0_4px_20px_rgba(0,0,0,0.15)]">
              <div className="flex items-center justify-between">
                <span className="text-xs text-slate-400 uppercase tracking-widest">Pending Tribunal</span>
                <span className="w-7 h-7 bg-amber-500/10 text-amber-400 rounded-lg flex items-center justify-center border border-amber-500/30">⏳</span>
              </div>
              <div className="space-y-1">
                <h3 className="text-3xl font-black text-amber-400 leading-none tracking-tight">{pendingCount.toLocaleString()}</h3>
                <p className="text-[10px] text-slate-400">{((pendingCount / (totalCount || 1)) * 100).toFixed(1)}% of filtered total</p>
              </div>
            </div>

            {/* Under investigation */}
            <div className="glass-panel glass-panel-hover rounded-xl p-5 relative overflow-hidden flex flex-col justify-between h-32 shadow-[0_4px_20px_rgba(0,0,0,0.15)]">
              <div className="flex items-center justify-between">
                <span className="text-xs text-slate-400 uppercase tracking-widest">Under Inquiry</span>
                <span className="w-7 h-7 bg-indigo-500/10 text-indigo-400 rounded-lg flex items-center justify-center border border-indigo-500/30">⚖️</span>
              </div>
              <div className="space-y-1">
                <h3 className="text-3xl font-black text-indigo-400 leading-none tracking-tight">{investigationCount.toLocaleString()}</h3>
                <p className="text-[10px] text-slate-400">{((investigationCount / (totalCount || 1)) * 100).toFixed(1)}% of filtered total</p>
              </div>
            </div>

            {/* Resolved complaints */}
            <div className="glass-panel glass-panel-hover rounded-xl p-5 relative overflow-hidden flex flex-col justify-between h-32 shadow-[0_4px_20px_rgba(0,0,0,0.15)]">
              <div className="flex items-center justify-between">
                <span className="text-xs text-slate-400 uppercase tracking-widest">Disposed/Resolved</span>
                <span className="w-7 h-7 bg-emerald-500/10 text-emerald-400 rounded-lg flex items-center justify-center border border-emerald-500/30">✓</span>
              </div>
              <div className="space-y-1">
                <h3 className="text-3xl font-black text-emerald-400 leading-none tracking-tight">{resolvedCount.toLocaleString()}</h3>
                <p className="text-[10px] text-slate-400">Case closure rate: {resolutionRate}%</p>
              </div>
            </div>
          </section>

          {/* Interactive Dynamic Charts Grid */}
          <section className="grid grid-cols-1 lg:grid-cols-5 gap-6">
            
            {/* Top 5 Crime Categories (3 Cols) */}
            <div className="lg:col-span-3 glass-panel rounded-2xl p-6 flex flex-col justify-between space-y-6 shadow-xl border border-police-700/20">
              <div className="flex items-center justify-between pb-3 border-b border-police-700/20">
                <div>
                  <h3 className="text-sm font-bold text-white uppercase tracking-wider">Crime Type distribution</h3>
                  <p className="text-[11px] text-slate-400">Top 5 categories by frequency in selection</p>
                </div>
                <span className="text-[10px] text-blue-400 bg-blue-950 px-2 py-0.5 border border-blue-900/50 rounded font-semibold">Live Analytics</span>
              </div>

              {categoryChartData.length === 0 ? (
                <div className="flex-1 flex items-center justify-center text-slate-500 text-xs h-40">
                  NO ACCESSIBLE INCIDENT RECORDS FOR CURRENT FILTER SETTINGS
                </div>
              ) : (
                <div className="space-y-4">
                  {categoryChartData.map((cat) => {
                    const maxCount = categoryChartData[0]?.count || 1;
                    const percentage = (cat.count / maxCount) * 100;
                    
                    return (
                      <div key={cat.name} className="space-y-1">
                        <div className="flex justify-between text-xs">
                          <span className="text-slate-300 font-semibold truncate max-w-[280px]" title={cat.name}>
                            {cat.name}
                          </span>
                          <span className="text-blue-400 font-bold">{cat.count.toLocaleString()} cases</span>
                        </div>
                        <div className="h-2.5 bg-[#05070f] rounded-full overflow-hidden border border-police-700/40">
                          <div 
                            className="h-full bg-gradient-to-r from-blue-600 to-indigo-400 rounded-full shadow-[0_0_8px_rgba(59,130,246,0.4)] transition-all duration-700 ease-out"
                            style={{ width: `${percentage}%` }}
                          />
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            {/* Resolution Rate & Reception Mode (2 Cols) */}
            <div className="lg:col-span-2 glass-panel rounded-2xl p-6 flex flex-col justify-between space-y-6 shadow-xl border border-police-700/20">
              <div className="flex items-center justify-between pb-3 border-b border-police-700/20">
                <div>
                  <h3 className="text-sm font-bold text-white uppercase tracking-wider">Metrics index</h3>
                  <p className="text-[11px] text-slate-400">Resolution index and submission sources</p>
                </div>
                <span className="text-[10px] text-emerald-400 bg-emerald-950 px-2 py-0.5 border border-emerald-900/50 rounded font-semibold">Disposal Index</span>
              </div>

              <div className="flex flex-col sm:flex-row items-center justify-around gap-6 w-full">
                
                {/* Circular glowing resolution progress */}
                <div className="flex flex-col items-center space-y-2 shrink-0">
                  <div className="relative flex items-center justify-center">
                    <svg className="w-28 h-28 transform -rotate-90">
                      <circle
                        className="text-police-800"
                        strokeWidth="7"
                        stroke="currentColor"
                        fill="transparent"
                        r={ringRadius}
                        cx="56"
                        cy="56"
                      />
                      <circle
                        className="text-emerald-500 transition-all duration-1000 ease-out filter drop-shadow-[0_0_6px_rgba(16,185,129,0.5)]"
                        strokeWidth="7"
                        strokeDasharray={ringCircumference}
                        strokeDashoffset={ringOffset}
                        strokeLinecap="round"
                        stroke="currentColor"
                        fill="transparent"
                        r={ringRadius}
                        cx="56"
                        cy="56"
                      />
                    </svg>
                    <div className="absolute flex flex-col items-center justify-center">
                      <span className="text-lg font-black text-white leading-none">{resolutionRate}%</span>
                      <span className="text-[8px] text-slate-400 uppercase tracking-widest mt-1">DISPOSED</span>
                    </div>
                  </div>
                </div>

                {/* Vertical lists for submission modes */}
                <div className="flex-1 space-y-2.5 w-full sm:max-w-[200px]">
                  <span className="text-[10px] text-slate-400 uppercase tracking-wider block border-b border-police-700/20 pb-1 font-semibold">Ingress Source</span>
                  {submissionModeData.length === 0 ? (
                    <p className="text-[10px] text-slate-500">No source records available</p>
                  ) : (
                    submissionModeData.map((mode) => {
                      const maxModeCount = submissionModeData[0]?.count || 1;
                      const modePercentage = (mode.count / maxModeCount) * 100;
                      return (
                        <div key={mode.name} className="space-y-0.5">
                          <div className="flex justify-between text-[10px] text-slate-300">
                            <span className="truncate max-w-[120px]">{mode.name}</span>
                            <span className="font-semibold text-slate-400">{mode.count}</span>
                          </div>
                          <div className="h-1 bg-[#05070f] rounded-full overflow-hidden">
                            <div 
                              className="h-full bg-gradient-to-r from-emerald-500 to-teal-400 rounded-full" 
                              style={{ width: `${modePercentage}%` }}
                            />
                          </div>
                        </div>
                      );
                    })
                  )}
                </div>

              </div>
            </div>
          </section>

          {/* Interactive Database Console Table */}
          <section className="glass-panel rounded-2xl shadow-2xl border border-police-700/20 overflow-hidden">
            
            {/* Table Control Panel */}
            <div className="px-6 py-5 border-b border-police-700/20 bg-[#070b1b]/80 flex flex-col md:flex-row md:items-center justify-between gap-4">
              
              <div className="flex items-center space-x-3.5">
                <h2 className="text-md font-bold text-white uppercase tracking-wider flex items-center">
                  Registry Database
                </h2>
                <span className="px-2.5 py-0.5 text-[10px] font-semibold bg-blue-950 text-blue-400 border border-blue-900/50 rounded-full">
                  Ingested {totalCount.toLocaleString()} cases
                </span>
              </div>

              {/* Table search & selection filters */}
              <div className="flex flex-col sm:flex-row flex-wrap items-stretch sm:items-center gap-3 w-full md:w-auto">
                
                {/* Search Bar */}
                <div className="relative w-full sm:w-64">
                  <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none">
                    <SearchIcon />
                  </div>
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => { setSearchQuery(e.target.value); setCurrentPage(1); }}
                    placeholder="Search ID, Name, Mobile, Description..."
                    className="pl-9 pr-4 py-2 w-full bg-[#05070f] border border-police-700/50 rounded-lg text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-blue-500/50 transition-all"
                  />
                </div>

                {/* Category Dropdown Filter */}
                <select
                  value={selectedCategory}
                  onChange={(e) => { setSelectedCategory(e.target.value); setCurrentPage(1); }}
                  className="bg-[#05070f] border border-police-700/50 rounded-lg px-3 py-2 text-xs text-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-blue-500/50 transition-all cursor-pointer w-full sm:w-44"
                >
                  <option value="">[ All Incident Types ]</option>
                  {uniqueCategories.map(cat => (
                    <option key={cat} value={cat}>{cat.toUpperCase()}</option>
                  ))}
                </select>

                {/* Export PDF Button */}
                <button
                  onClick={exportPDF}
                  className="bg-red-950/60 hover:bg-red-900/80 border border-red-800/60 hover:border-red-600/80 px-3.5 py-2 text-xs font-bold text-red-400 hover:text-white rounded-lg transition-all cursor-pointer flex items-center justify-center gap-1.5 shadow-[0_0_10px_rgba(239,68,68,0.1)] hover:shadow-[0_0_15px_rgba(239,68,68,0.25)] w-full sm:w-auto"
                  title="Export up to 1000 filtered rows to PDF"
                >
                  <span>📄</span> Export PDF
                </button>

              </div>
            </div>

            {/* The Database Grid */}
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-police-700/30">
                <thead className="bg-[#050711]">
                  <tr>
                    
                    <th onClick={() => toggleSort("id")} className="px-6 py-4 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-widest cursor-pointer hover:bg-police-950 transition-colors">
                      <div className="flex items-center space-x-1.5">
                        <span>Case File ID</span>
                        <SortIcon active={sortConfig?.key === "id"} direction={sortConfig?.direction || "desc"} />
                      </div>
                    </th>

                    <th onClick={() => toggleSort("date")} className="px-6 py-4 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-widest cursor-pointer hover:bg-police-950 transition-colors hidden sm:table-cell">
                      <div className="flex items-center space-x-1.5">
                        <span>Filing Date</span>
                        <SortIcon active={sortConfig?.key === "date"} direction={sortConfig?.direction || "desc"} />
                      </div>
                    </th>

                    <th onClick={() => toggleSort("district")} className="px-6 py-4 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-widest cursor-pointer hover:bg-police-950 transition-colors hidden md:table-cell">
                      <div className="flex items-center space-x-1.5">
                        <span>Geographic District</span>
                        <SortIcon active={sortConfig?.key === "district"} direction={sortConfig?.direction || "desc"} />
                      </div>
                    </th>

                    <th className="px-6 py-4 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-widest hidden lg:table-cell">Office Master</th>
                    
                    <th className="px-6 py-4 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-widest hidden lg:table-cell">Police Station</th>
                    
                    <th onClick={() => toggleSort("status")} className="px-6 py-4 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-widest cursor-pointer hover:bg-police-950 transition-colors">
                      <div className="flex items-center space-x-1.5">
                        <span>Consolidated Status</span>
                        <SortIcon active={sortConfig?.key === "status"} direction={sortConfig?.direction || "desc"} />
                      </div>
                    </th>

                    <th className="px-6 py-4 text-left text-[11px] font-semibold text-slate-400 uppercase tracking-widest hidden md:table-cell">Complainant</th>
                    
                    <th className="px-6 py-4 text-center text-[11px] font-semibold text-slate-400 uppercase tracking-widest">Action</th>

                  </tr>
                </thead>
                <tbody className="divide-y divide-police-700/20 bg-transparent text-xs text-slate-300">
                  {tableData.length === 0 ? (
                    <tr>
                      <td colSpan={8} className="px-6 py-12 text-center text-slate-500">
                        NO RECORDS MATCHING THE DEPLOYED FILTERING PARAMETERS
                      </td>
                    </tr>
                  ) : (
                    tableData.map((complaint, index) => (
                      <tr 
                        key={`${complaint.id}-${index}`} 
                        className={`hover:bg-police-800/25 transition-all group ${
                          selectedComplaint && selectedComplaint.complRegNum === complaint.raw.complRegNum 
                            ? "bg-blue-600/5 border-l-2 border-l-blue-500" 
                            : ""
                        }`}
                      >
                        <td className="px-6 py-4 whitespace-nowrap text-slate-100 font-bold tracking-wider">{complaint.id}</td>
                        <td className="px-6 py-4 whitespace-nowrap text-slate-300 hidden sm:table-cell">{complaint.date}</td>
                        <td className="px-6 py-4 whitespace-nowrap text-slate-300 hidden md:table-cell">{complaint.district}</td>
                        <td className="px-6 py-4 whitespace-nowrap text-slate-400 truncate max-w-[150px] hidden lg:table-cell" title={complaint.office}>{complaint.office}</td>
                        <td className="px-6 py-4 whitespace-nowrap text-slate-400 truncate max-w-[150px] hidden lg:table-cell" title={complaint.policeStation}>{complaint.policeStation}</td>
                        
                        <td className="px-6 py-4 whitespace-nowrap">
                          <span className={`px-2.5 py-0.5 rounded-full text-[10px] font-bold ${
                            complaint.status.toLowerCase() === "pending" ? "bg-amber-500/10 text-amber-400 border border-amber-500/20" :
                            complaint.status.toLowerCase() === "disposed" ? "bg-emerald-500/10 text-emerald-400 border border-emerald-500/20" :
                            complaint.status.toLowerCase().includes("investigation") ? "bg-indigo-500/10 text-indigo-400 border border-indigo-500/20" :
                            "bg-slate-800 text-slate-300 border border-slate-700/50"
                          }`}>
                            {complaint.status.toUpperCase()}
                          </span>
                        </td>

                        <td className="px-6 py-4 whitespace-nowrap text-slate-200 hidden md:table-cell">{complaint.complainant}</td>
                        
                        <td className="px-6 py-4 whitespace-nowrap text-center">
                          <button
                            onClick={() => setSelectedComplaint(complaint.raw)}
                            className="px-3.5 py-1 text-[11px] font-bold bg-blue-600/10 hover:bg-blue-600/80 border border-blue-500/35 hover:border-blue-500 text-blue-400 hover:text-white rounded-md transition-all shadow-[0_0_6px_rgba(59,130,246,0.1)] cursor-pointer group-hover:scale-105"
                          >
                            Inspect
                          </button>
                        </td>

                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination Controls */}
            {totalPages > 1 && (
              <div className="px-6 py-4 border-t border-police-700/20 bg-[#070b1b]/50 flex items-center justify-between text-xs">
                <div className="text-slate-400">
                  PAGE <span className="text-white font-bold">{currentPage}</span> OF <span className="text-white font-bold">{totalPages}</span>
                </div>
                <div className="flex space-x-2">
                  <button
                    onClick={() => { setCurrentPage(p => Math.max(1, p - 1)); }}
                    disabled={currentPage === 1}
                    className="px-3 py-1.5 bg-[#05070f] border border-police-700/40 rounded-lg text-slate-300 hover:text-white hover:border-blue-500 disabled:opacity-20 disabled:cursor-not-allowed disabled:hover:border-police-700/40 disabled:hover:text-slate-300 transition-all cursor-pointer"
                  >
                    PREV
                  </button>
                  <button
                    onClick={() => { setCurrentPage(p => Math.min(totalPages, p + 1)); }}
                    disabled={currentPage === totalPages}
                    className="px-3 py-1.5 bg-[#05070f] border border-police-700/40 rounded-lg text-slate-300 hover:text-white hover:border-blue-500 disabled:opacity-20 disabled:cursor-not-allowed disabled:hover:border-police-700/40 disabled:hover:text-slate-300 transition-all cursor-pointer"
                  >
                    NEXT
                  </button>
                </div>
              </div>
            )}

          </section>
        </div>
      </main>

      {/* Slide-over Inspection Panel / Drawer Overlay */}
      {selectedComplaint && (
        <div className="fixed inset-0 z-50 flex justify-end">
          
          {/* Tinted backdrop */}
          <div 
            onClick={() => setSelectedComplaint(null)}
            className="absolute inset-0 bg-[#04060f]/60 backdrop-blur-sm animate-fade-in"
          />

          {/* Drawer Body Container */}
          <div className="w-full max-w-[620px] bg-[#070b1b] border-l border-police-700/35 h-full relative z-10 flex flex-col shadow-[0_0_50px_rgba(0,0,0,0.8)] animate-slide-in">
            
            {/* Header banner */}
            <div className="p-4 sm:p-6 border-b border-police-700/30 bg-gradient-to-r from-police-950 to-police-900 flex items-center justify-between">
              <div>
                <span className="text-[10px] text-blue-400 uppercase tracking-widest">Case Inquest Dossier</span>
                <h3 className="text-md font-extrabold text-white tracking-wider flex items-center mt-1">
                  ID: {(() => {
                    const raw = String(selectedComplaint.complRegNum || "").trim();
                    try {
                      if (raw.includes('E') || raw.includes('e')) {
                        const num = Number(raw);
                        return !isNaN(num) ? BigInt(Math.round(num)).toString() : raw;
                      }
                      return raw;
                    } catch {
                      return raw;
                    }
                  })()}
                </h3>
              </div>
              <button 
                onClick={() => setSelectedComplaint(null)}
                className="p-2 bg-police-950 rounded-lg hover:bg-police-800 border border-police-700/40 cursor-pointer transition-all"
              >
                <CloseIcon />
              </button>
            </div>

            {/* Dossier Body Content */}
            <div className="flex-1 overflow-y-auto p-4 sm:p-6 space-y-6">
              
              {/* Status Section */}
              <div className="glass-panel rounded-xl p-4 md:p-4.5 space-y-3 bg-[#0d142a]/30">
                <div className="flex items-center justify-between border-b border-police-700/20 pb-2">
                  <span className="text-[10px] text-slate-400 uppercase tracking-wider">Investigative Status</span>
                  <span className={`px-2.5 py-0.5 rounded-full text-[10px] font-bold ${
                    String(selectedComplaint.statusGroup || "").toLowerCase() === "pending" ? "bg-amber-500/10 text-amber-400 border border-amber-500/20" :
                    String(selectedComplaint.statusGroup || "").toLowerCase() === "disposed" ? "bg-emerald-500/10 text-emerald-400 border border-emerald-500/20" :
                    String(selectedComplaint.statusGroup || "").toLowerCase().includes("investigation") ? "bg-indigo-500/10 text-indigo-400 border border-indigo-500/20" :
                    "bg-slate-800 text-slate-300 border border-slate-700/50"
                  }`}>
                    {String(selectedComplaint.statusGroup || selectedComplaint.statusOfComplaint || selectedComplaint.statusRaw || "UNKNOWN").toUpperCase()}
                  </span>
                </div>
                
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
                  <div>
                    <p className="text-[10px] text-slate-500 uppercase">Registration Date</p>
                    <p className="text-slate-200 font-bold mt-0.5">{parseDate(selectedComplaint.complRegDt)}</p>
                  </div>
                  <div>
                    <p className="text-[10px] text-slate-500 uppercase">Disposal Date</p>
                    <p className={`font-bold mt-0.5 ${selectedComplaint.disposalDate ? "text-emerald-400" : "text-slate-400"}`}>
                      {selectedComplaint.disposalDate ? parseDate(selectedComplaint.disposalDate) : "Awaiting Resolution"}
                    </p>
                  </div>
                </div>
              </div>

              {/* Complainant Identity Section */}
              <div className="space-y-3">
                <h4 className="text-xs font-semibold tracking-wider text-slate-400 uppercase flex items-center">
                  <span className="w-1.5 h-1.5 bg-blue-500 rounded-full mr-2"></span>
                  Complainant Identity
                </h4>
                
                <div className="glass-panel rounded-xl p-4 md:p-4.5 grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs bg-[#0d142a]/30">
                  <div className="col-span-1 sm:col-span-2">
                    <p className="text-[10px] text-slate-500 uppercase">Complainant Full Name</p>
                    <p className="text-slate-200 font-bold text-sm mt-0.5">
                      {`${selectedComplaint.firstName || ""} ${selectedComplaint.lastName || ""}`.trim() || "N/A (ANONYMOUS/SECURITY REGISTRY)"}
                    </p>
                  </div>
                  
                  <div>
                    <p className="text-[10px] text-slate-500 uppercase">Biological Profile</p>
                    <p className="text-slate-200 font-bold mt-0.5">
                      {selectedComplaint.gender ? selectedComplaint.gender.toUpperCase() : "N/A"}
                      {selectedComplaint.age ? `, Age: ${selectedComplaint.age}` : ""}
                    </p>
                  </div>

                  <div>
                    <p className="text-[10px] text-slate-500 uppercase">Contact Link</p>
                    <p className="text-slate-200 font-bold mt-0.5">{selectedComplaint.mobile || "N/A"}</p>
                  </div>

                  {selectedComplaint.email && (
                    <div className="col-span-1 sm:col-span-2">
                      <p className="text-[10px] text-slate-500 uppercase">Email Dispatch Link</p>
                      <p className="text-slate-200 font-bold mt-0.5">{selectedComplaint.email}</p>
                    </div>
                  )}

                  <div className="col-span-1 sm:col-span-2 border-t border-police-700/20 pt-3 space-y-2">
                    <p className="text-[10px] text-slate-500 uppercase">Residential Address</p>
                    <p className="text-slate-300 font-medium leading-relaxed">
                      {[
                        selectedComplaint.addressLine1,
                        selectedComplaint.addressLine2,
                        selectedComplaint.addressLine3,
                        selectedComplaint.village,
                        selectedComplaint.tehsil,
                        selectedComplaint.addressDistrict,
                        selectedComplaint.addressPs
                      ].filter(Boolean).join(", ") || "No address log on file"}
                    </p>
                  </div>
                </div>
              </div>

              {/* Incident Details Section */}
              <div className="space-y-3">
                <h4 className="text-xs font-semibold tracking-wider text-slate-400 uppercase flex items-center">
                  <span className="w-1.5 h-1.5 bg-indigo-500 rounded-full mr-2"></span>
                  Incident Case Details
                </h4>
                
                <div className="glass-panel rounded-xl p-4 md:p-4.5 space-y-4 text-xs bg-[#0d142a]/30">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <p className="text-[10px] text-slate-500 uppercase">Incident Type / Classification</p>
                      <p className="text-blue-400 font-bold mt-0.5 uppercase">{selectedComplaint.incidentType || "Unknown Type"}</p>
                    </div>
                    <div>
                      <p className="text-[10px] text-slate-500 uppercase">Incident Scene Location</p>
                      <p className="text-slate-200 font-bold mt-0.5">{selectedComplaint.incidentPlc || "Not Logged"}</p>
                    </div>
                    <div>
                      <p className="text-[10px] text-slate-500 uppercase">Occurred From</p>
                      <p className="text-slate-200 font-bold mt-0.5">{selectedComplaint.incidentFromDt ? parseDate(selectedComplaint.incidentFromDt) : "N/A"}</p>
                    </div>
                    <div>
                      <p className="text-[10px] text-slate-500 uppercase">Occurred To</p>
                      <p className="text-slate-200 font-bold mt-0.5">{selectedComplaint.incidentToDt ? parseDate(selectedComplaint.incidentToDt) : "N/A"}</p>
                    </div>
                  </div>

                  <div className="border-t border-police-700/20 pt-3">
                    <p className="text-[10px] text-slate-500 uppercase mb-2">Complainant Narrative Description</p>
                    <div className="bg-[#05070f] border border-police-700/50 p-4.5 rounded-lg text-slate-300 leading-relaxed text-xs overflow-y-auto max-h-48 whitespace-pre-wrap">
                      {selectedComplaint.complDesc || "No narrative description provided in registration file."}
                    </div>
                  </div>
                </div>
              </div>

              {/* Action Log / IO details */}
              <div className="space-y-3">
                <h4 className="text-xs font-semibold tracking-wider text-slate-400 uppercase flex items-center">
                  <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full mr-2"></span>
                  HQ Operational Assignment
                </h4>
                
                <div className="glass-panel rounded-xl p-4 md:p-4.5 grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs bg-[#0d142a]/30">
                  <div className="col-span-1 sm:col-span-2">
                    <p className="text-[10px] text-slate-500 uppercase">Investigating Officer (IO Details)</p>
                    <p className="text-slate-200 font-bold mt-1 text-[13px]">
                      {selectedComplaint.ioDetails || "No Investigating Officer assigned yet."}
                    </p>
                  </div>
                  <div>
                    <p className="text-[10px] text-slate-500 uppercase">Complaint Source Channel</p>
                    <p className="text-slate-300 font-bold mt-0.5">{selectedComplaint.complaintSource || "Direct System Ingress"}</p>
                  </div>
                  <div>
                    <p className="text-[10px] text-slate-500 uppercase">Purpose Class</p>
                    <p className="text-slate-300 font-bold mt-0.5">{selectedComplaint.complaintPurpose || "N/A"}</p>
                  </div>
                </div>
              </div>

            </div>
          </div>
        </div>
      )}

      {/* Import Status Toast Alert */}
      {importStatus && (
        <div className="fixed bottom-5 right-5 z-50 bg-[#0d1530] border border-emerald-500/50 shadow-[0_0_15px_rgba(16,185,129,0.3)] px-5 py-3 rounded-lg text-emerald-400 text-xs font-bold animate-pulse flex items-center space-x-2">
          <span className="w-2 h-2 bg-emerald-400 rounded-full animate-ping"></span>
          <span>{importStatus}</span>
        </div>
      )}

    </div>
  );
}