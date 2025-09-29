"use client";

import { useState, useEffect, useContext } from "react";
import { MapPin, Loader2, Download } from "lucide-react";
import { AuthContext } from "../App";

// AttendanceHistory Component with filters in header and Excel download functionality
const AttendanceHistory = ({ attendanceData, isLoading, userRole }) => {
  const [filters, setFilters] = useState({
    name: "",
    status: "",
    month: ""
  });
  const [filteredData, setFilteredData] = useState([]);

  // Month names for dropdown
  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  const getUniqueNames = (data) => {
    const names = data.map(entry => entry.salesPersonName).filter(Boolean);
    return [...new Set(names)].sort();
  };

  const getAvailableMonths = (data) => {
    const months = new Set();
    data.forEach(entry => {
      if (entry.dateTime) {
        const dateStr = entry.dateTime.split(" ")[0];
        if (dateStr) {
          const [day, month, year] = dateStr.split("/");
          const monthNum = parseInt(month, 10) - 1;
          const monthName = monthNames[monthNum];
          if (monthName) {
            months.add(`${monthName} ${year}`);
          }
        }
      }
    });
    return Array.from(months).sort((a, b) => {
      const [aMonth, aYear] = a.split(" ");
      const [bMonth, bYear] = b.split(" ");
      const aDate = new Date(parseInt(aYear), monthNames.indexOf(aMonth));
      const bDate = new Date(parseInt(bYear), monthNames.indexOf(bMonth));
      return bDate - aDate;
    });
  };

  const applyFilters = (data) => {
    if (!filters.name && !filters.status && !filters.month) {
      return data;
    }

    return data.filter((entry) => {
      if (filters.name && !entry.salesPersonName?.toLowerCase().includes(filters.name.toLowerCase())) {
        return false;
      }

      if (filters.status && entry.status !== filters.status) {
        return false;
      }

      if (filters.month) {
        const entryDate = entry.dateTime?.split(" ")[0];
        if (entryDate) {
          const [day, month, year] = entryDate.split("/");
          const monthNum = parseInt(month, 10) - 1;
          const monthName = monthNames[monthNum];
          const entryMonthYear = `${monthName} ${year}`;
          if (entryMonthYear !== filters.month) {
            return false;
          }
        }
      }

      return true;
    });
  };

  // Enhanced Excel download function
  const downloadExcel = () => {
    if (!filteredData || filteredData.length === 0) {
      alert('No data available to download');
      return;
    }

    // Create proper Excel content with XML format
    const currentDate = new Date().toLocaleDateString();
    const fileName = `Attendance_History_${new Date().toISOString().split('T')[0]}`;
    
    // Create Excel XML structure
    let excelContent = `<?xml version="1.0"?>
      <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
                xmlns:o="urn:schemas-microsoft-com:office:office"
                xmlns:x="urn:schemas-microsoft-com:office:excel"
                xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
                xmlns:html="http://www.w3.org/TR/REC-html40">
        <Worksheet ss:Name="Attendance History">
          <Table>
            <Row>
              <Cell><Data ss:Type="String">Name</Data></Cell>
              <Cell><Data ss:Type="String">Date &amp; Time</Data></Cell>
              <Cell><Data ss:Type="String">Status</Data></Cell>
              <Cell><Data ss:Type="String">Map Link</Data></Cell>
              <Cell><Data ss:Type="String">Address</Data></Cell>
            </Row>`;

    // Add data rows
    filteredData.forEach(row => {
      excelContent += `
        <Row>
          <Cell><Data ss:Type="String">${row.salesPersonName || 'N/A'}</Data></Cell>
          <Cell><Data ss:Type="String">${row.dateTime || 'N/A'}</Data></Cell>
          <Cell><Data ss:Type="String">${row.status || 'N/A'}</Data></Cell>
          <Cell><Data ss:Type="String">${row.mapLink || 'N/A'}</Data></Cell>
          <Cell><Data ss:Type="String">${(row.address || 'N/A').replace(/[<>&"']/g, function(match) {
            switch(match) {
              case '<': return '&lt;';
              case '>': return '&gt;';
              case '&': return '&amp;';
              case '"': return '&quot;';
              case "'": return '&apos;';
              default: return match;
            }
          })}</Data></Cell>
        </Row>`;
    });

    excelContent += `
          </Table>
        </Worksheet>
      </Workbook>`;

    // Create and download the file
    const blob = new Blob([excelContent], { 
      type: 'application/vnd.ms-excel;charset=utf-8;' 
    });
    
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', `${fileName}.xls`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  useEffect(() => {
    const filtered = applyFilters(attendanceData || []);
    setFilteredData(filtered);
  }, [filters, attendanceData]);

  const handleFilterChange = (filterType, value) => {
    setFilters(prev => ({
      ...prev,
      [filterType]: value
    }));
  };

  const clearFilters = () => {
    setFilters({
      name: "",
      status: "",
      month: ""
    });
  };

  const hasActiveFilters = filters.name || filters.status || filters.month;

  if (isLoading) {
    return (
      <div className="bg-white/80 backdrop-blur-sm rounded-2xl shadow-xl border border-white/20 overflow-hidden mt-8">
        <div className="bg-gradient-to-r from-blue-500 via-indigo-500 to-purple-500 px-8 py-6">
          <h2 className="text-2xl font-bold text-white">Attendance History</h2>
          <p className="text-blue-50">Loading your attendance records...</p>
        </div>
        <div className="p-8 text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-indigo-600 mx-auto mb-4"></div>
          <p className="text-slate-600">Loading attendance history...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="bg-white/80 backdrop-blur-sm rounded-2xl shadow-xl border border-white/20 overflow-hidden mt-8">
      {/* Header with Filters and Download */}
      <div className="bg-gradient-to-r from-blue-500 via-indigo-500 to-purple-500 px-8 py-6">
        <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4 mb-4">
          <div>
            <h2 className="text-2xl font-bold text-white">Attendance History</h2>
            <p className="text-blue-50">Your records are displayed below.</p>
          </div>
          
          {/* Excel Download Button Only */}
          {userRole?.toLowerCase() === "admin" && filteredData.length > 0 && (
            <div className="flex gap-2">
              <button
                onClick={downloadExcel}
                className="flex items-center gap-2 px-4 py-2 bg-green-600 hover:bg-green-700 text-white text-sm rounded-lg border border-green-500 transition-colors shadow-md"
                title="Download as Excel"
              >
                <Download className="h-4 w-4" />
                Download 
              </button>
            </div>
          )}
        </div>

        {/* Filters Row - Only show for admin */}
        {userRole?.toLowerCase() === "admin" && (
          <div className="grid gap-3 md:grid-cols-4 items-end">
            {/* Name Filter */}
            <div>
              <label className="block text-sm font-medium text-blue-100 mb-1">
                Filter by Name
              </label>
              <select
                value={filters.name}
                onChange={(e) => handleFilterChange('name', e.target.value)}
                className="w-full px-3 py-2 bg-white/90 border border-white/30 rounded-lg text-slate-700 text-sm focus:ring-2 focus:ring-white/50 focus:border-white/50"
              >
                <option value="">All Names</option>
                {getUniqueNames(attendanceData || []).map((name) => (
                  <option key={name} value={name}>
                    {name}
                  </option>
                ))}
              </select>
            </div>

            {/* Status Filter */}
            <div>
              <label className="block text-sm font-medium text-blue-100 mb-1">
                Filter by Status
              </label>
              <select
                value={filters.status}
                onChange={(e) => handleFilterChange('status', e.target.value)}
                className="w-full px-3 py-2 bg-white/90 border border-white/30 rounded-lg text-slate-700 text-sm focus:ring-2 focus:ring-white/50 focus:border-white/50"
              >
                <option value="">All Status</option>
                <option value="IN">IN</option>
                <option value="OUT">OUT</option>
                <option value="Leave">Leave</option>
              </select>
            </div>

            {/* Month Filter */}
            <div>
              <label className="block text-sm font-medium text-blue-100 mb-1">
                Filter by Month
              </label>
              <select
                value={filters.month}
                onChange={(e) => handleFilterChange('month', e.target.value)}
                className="w-full px-3 py-2 bg-white/90 border border-white/30 rounded-lg text-slate-700 text-sm focus:ring-2 focus:ring-white/50 focus:border-white/50"
              >
                <option value="">All Months</option>
                {getAvailableMonths(attendanceData || []).map((monthYear) => (
                  <option key={monthYear} value={monthYear}>
                    {monthYear}
                  </option>
                ))}
              </select>
            </div>

            {/* Clear Filters Button */}
            <div>
              {hasActiveFilters && (
                <button
                  onClick={clearFilters}
                  className="w-full px-3 py-2 bg-white/20 hover:bg-white/30 text-white text-sm rounded-lg border border-white/30 transition-colors"
                >
                  Clear Filters
                </button>
              )}
            </div>
          </div>
        )}

        {/* Filter Results Info */}
        {hasActiveFilters && (
          <div className="mt-3 bg-white/10 border border-white/20 rounded-lg p-3">
            <p className="text-sm text-blue-100">
              Showing {filteredData.length} of {attendanceData?.length || 0} records
              {filters.name && ` • Name: ${filters.name}`}
              {filters.status && ` • Status: ${filters.status}`}
              {filters.month && ` • Month: ${filters.month}`}
            </p>
          </div>
        )}
      </div>

      {/* Table Content */}
      <div className="overflow-x-auto">
        {(!attendanceData || attendanceData.length === 0) ? (
          <div className="p-8 text-center">
            <div className="text-slate-400 text-lg mb-2">📊</div>
            <h3 className="text-lg font-semibold text-slate-600 mb-2">
              No Records Found
            </h3>
            <p className="text-slate-500">
              {userRole?.toLowerCase() === "admin" 
                ? "No attendance records available."
                : "You haven't marked any attendance yet."}
            </p>
          </div>
        ) : filteredData.length === 0 && hasActiveFilters ? (
          <div className="p-8 text-center">
            <div className="text-slate-400 text-lg mb-2">🔍</div>
            <h3 className="text-lg font-semibold text-slate-600 mb-2">
              No Matching Records
            </h3>
            <p className="text-slate-500">
              No records match your current filter criteria.
            </p>
          </div>
        ) : (
          <div className="min-w-full">
            <table className="w-full border-collapse">
              <thead className="bg-slate-50/50 border-b border-slate-200/50">
                <tr>
                  <th className="px-4 py-3 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider border-r border-slate-200/50 w-32">
                    Name
                  </th>
                  <th className="px-4 py-3 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider border-r border-slate-200/50 w-40">
                    Date & Time
                  </th>
                  <th className="px-4 py-3 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider border-r border-slate-200/50 w-24">
                    Status
                  </th>
                  <th className="px-4 py-3 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider border-r border-slate-200/50 w-32">
                    Map Link
                  </th>
                  <th className="px-4 py-3 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">
                    Address
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200/50">
                {filteredData.map((record, index) => (
                  <tr key={index} className="hover:bg-slate-50/30 transition-colors border-b border-slate-200/30">
                    <td className="px-4 py-3 border-r border-slate-200/50 w-32">
                      <div className="text-sm font-medium text-slate-900 break-words">
                        {record.salesPersonName || "N/A"}
                      </div>
                    </td>
                    <td className="px-4 py-3 border-r border-slate-200/50 w-40">
                      <div className="text-sm text-slate-900 break-words">
                        {record.dateTime || "N/A"}
                      </div>
                    </td>
                    <td className="px-4 py-3 border-r border-slate-200/50 w-24">
                      <span className={`inline-flex px-2 py-1 text-xs font-semibold rounded-full ${
                        record.status === "IN"
                          ? "bg-green-100 text-green-800"
                          : record.status === "OUT"
                          ? "bg-red-100 text-red-800"
                          : record.status === "Leave"
                          ? "bg-yellow-100 text-yellow-800"
                          : "bg-gray-100 text-gray-800"
                      }`}>
                        {record.status || "N/A"}
                      </span>
                    </td>
                    <td className="px-4 py-3 border-r border-slate-200/50 w-32">
                      {record.mapLink ? (
                        <a
                          href={record.mapLink}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="text-blue-600 hover:text-blue-800 flex items-center gap-1 text-sm break-all"
                        >
                          <MapPin className="h-4 w-4 flex-shrink-0" />
                          <span className="truncate">View Map</span>
                        </a>
                      ) : (
                        <span className="text-slate-400 text-sm">N/A</span>
                      )}
                    </td>
                    <td className="px-4 py-3">
                      <div className="text-sm text-slate-600 break-words max-w-md" title={record.address}>
                        {record.address || "N/A"}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};

// Main Attendance Component (unchanged from previous version)
const Attendance = () => {
  const [attendance, setAttendance] = useState([]);
  const [historyAttendance, setHistoryAttendance] = useState([]);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isGettingLocation, setIsGettingLocation] = useState(false);
  const [hasCheckedInToday, setHasCheckedInToday] = useState(false);
  const [errors, setErrors] = useState({});
  const [locationData, setLocationData] = useState(null);
  const [isLoadingHistory, setIsLoadingHistory] = useState(true);
  const [hasActiveSession, setHasActiveSession] = useState(false);
  const [hasOutActiveSession, setHasOutActiveSession] = useState([]);
  const [inData, setInData] = useState({});
  const [outData, setOutData] = useState({});

  const { currentUser, isAuthenticated } = useContext(AuthContext);

  const salesPersonName = currentUser?.salesPersonName || "Unknown User";
  const userRole = currentUser?.role || "User";

  const SPREADSHEET_ID = "1WTT8ZQhtf1yeSChNn2uJeW5Tz2TvYjQLrxhTx5l4Fgw";
  const APPS_SCRIPT_URL =
    "https://script.google.com/macros/s/AKfycbxwve2gvQqFeo_OAkIBVS5uzKX92fZJAEyYtgE0GWQPlxs-3r-ofYA00_mEM19LumWIUg/exec";

  const formatDateInput = (date) => {
    return date.toISOString().split("T")[0];
  };

  const formatDateMMDDYYYY = (date) => {
    const day = String(date.getDate()).padStart(2, "0");
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const year = date.getFullYear();
    return `${month}/${day}/${year}`;
  };

  const formatDateDDMMYYYY = (date) => {
    const day = String(date.getDate()).padStart(2, "0");
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  };

  const formatDateTime = (date) => {
    const day = String(date.getDate()).padStart(2, "0");
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const year = date.getFullYear();
    const hours = String(date.getHours()).padStart(2, "0");
    const minutes = String(date.getMinutes()).padStart(2, "0");
    const seconds = String(date.getSeconds()).padStart(2, "0");
    return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
  };

  const formatDateDisplay = (date) => {
    return date.toLocaleDateString("en-GB", {
      weekday: "long",
      year: "numeric",
      month: "long",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: true,
    });
  };

  const [formData, setFormData] = useState({
    status: "",
    startDate: formatDateInput(new Date()),
    endDate: "",
    reason: "",
  });

  const showToast = (message, type = "success") => {
    const toast = document.createElement("div");
    const bgColor = type === "error" ? "bg-red-500" : "bg-green-500";

    toast.className = `fixed top-4 right-4 p-4 rounded-md text-white z-50 ${bgColor} max-w-sm shadow-lg`;
    toast.textContent = message;
    document.body.appendChild(toast);

    setTimeout(() => {
      if (toast.parentNode) {
        document.body.removeChild(toast);
      }
    }, 3000);
  };

  const getFormattedAddress = async (latitude, longitude) => {
    try {
      const response = await fetch(
        `https://nominatim.openstreetmap.org/reverse?format=json&lat=${latitude}&lon=${longitude}&addressdetails=1`
      );
      const data = await response.json();

      if (data && data.display_name) {
        return data.display_name;
      } else {
        return `${latitude.toFixed(6)}, ${longitude.toFixed(6)}`;
      }
    } catch (error) {
      console.error("Error getting formatted address:", error);
      return `${latitude.toFixed(6)}, ${longitude.toFixed(6)}`;
    }
  };

  const getCurrentLocation = () => {
    return new Promise((resolve, reject) => {
      if (!navigator.geolocation) {
        reject(new Error("Geolocation is not supported by this browser."));
        return;
      }

      const options = {
        enableHighAccuracy: true,
        timeout: 15000,
        maximumAge: 0,
      };

      navigator.geolocation.getCurrentPosition(
        async (position) => {
          const latitude = position.coords.latitude;
          const longitude = position.coords.longitude;
          const mapLink = `https://www.google.com/maps/search/?api=1&query=${latitude},${longitude}`;

          const formattedAddress = await getFormattedAddress(
            latitude,
            longitude
          );

          const locationInfo = {
            latitude,
            longitude,
            mapLink,
            formattedAddress,
            timestamp: new Date().toISOString(),
            accuracy: position.coords.accuracy,
          };

          resolve(locationInfo);
        },
        (error) => {
          const errorMessages = {
            1: "Location permission denied. Please enable location services.",
            2: "Location information unavailable.",
            3: "Location request timed out.",
          };
          reject(
            new Error(errorMessages[error.code] || "An unknown error occurred.")
          );
        },
        options
      );
    });
  };

  const checkActiveSession = (attendanceData) => {
    if (!attendanceData || attendanceData.length === 0) {
      setHasActiveSession(false);
      setHasCheckedInToday(false);
      return;
    }

    const userRecords = attendanceData.filter(
      (record) =>
        record.salesPersonName === salesPersonName &&
        record.dateTime?.split(" ")[0].toString() ===
          formatDateDDMMYYYY(new Date())
    );

    if (userRecords.length === 0) {
      setHasActiveSession(false);
      setHasCheckedInToday(false);
      return;
    }

    const mostRecentRecord = userRecords[0];
    const hasActive = mostRecentRecord.status === "IN";
    setHasActiveSession(hasActive);
    if (hasActive) {
      setInData(mostRecentRecord);
    }

    const hasOutActive = mostRecentRecord.status === "OUT";
    if (hasOutActive) {
      setOutData(mostRecentRecord);
    }

    const hasCheckedIn = userRecords.some(record => record.status === "IN");
    setHasCheckedInToday(hasCheckedIn);
  };

  const validateForm = () => {
    const newErrors = {};

    if (!formData.status) newErrors.status = "Status is required";

    if (formData.status === "Leave") {
      if (!formData.startDate) newErrors.startDate = "Start date is required";
      if (
        formData.startDate &&
        formData.endDate &&
        new Date(formData.endDate + "T00:00:00") <
          new Date(formData.startDate + "T00:00:00")
      ) {
        newErrors.endDate = "End date cannot be before start date";
      }
      if (!formData.reason) newErrors.reason = "Reason is required for leave";
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const fetchAttendanceHistory = async () => {
    if (!isAuthenticated || !currentUser) {
      console.log(
        "Not authenticated or currentUser not available. Skipping history fetch."
      );
      setIsLoadingHistory(false);
      return;
    }

    setIsLoadingHistory(true);
    try {
      const attendanceSheetUrl = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=Attendance`;
      const response = await fetch(attendanceSheetUrl);
      const text = await response.text();

      const jsonStart = text.indexOf("{");
      const jsonEnd = text.lastIndexOf("}") + 1;
      const jsonData = text.substring(jsonStart, jsonEnd);
      const data = JSON.parse(jsonData);

      if (!data?.table?.rows) {
        console.warn("No rows found in Attendance sheet.");
        setAttendance([]);
        setIsLoadingHistory(false);
        return;
      }

      const rows = data.table.rows;
      const formattedHistory = rows
        .map((row) => {
          const salesPerson = row.c?.[9]?.v;
          let dateTime = row.c?.[1]?.v;
          let originalTimestamp = row.c?.[0]?.v;

          if (
            typeof originalTimestamp === "string" &&
            originalTimestamp.startsWith("Date(") &&
            originalTimestamp.endsWith(")")
          ) {
            try {
              const dateParts = originalTimestamp
                .substring(5, originalTimestamp.length - 1)
                .split(",");
              const year = parseInt(dateParts[0], 10);
              const month = parseInt(dateParts[1], 10);
              const day = parseInt(dateParts[2], 10);
              const hour = dateParts[3] ? parseInt(dateParts[3], 10) : 0;
              const minute = dateParts[4] ? parseInt(dateParts[4], 10) : 0;
              const second = dateParts[5] ? parseInt(dateParts[5], 10) : 0;

              const dateObj = new Date(year, month, day, hour, minute, second);
              dateTime = formatDateTime(dateObj);
            } catch (e) {
              console.error(
                "Error parsing original timestamp date string:",
                originalTimestamp,
                e
              );
              dateTime = originalTimestamp;
            }
          }

          const status = row.c?.[3]?.v;
          const mapLink = row.c?.[7]?.v;
          const address = row.c?.[8]?.v;

          return {
            salesPersonName: salesPerson,
            dateTime: dateTime,
            status: status,
            mapLink: mapLink,
            address: address,
            _originalTimestamp: originalTimestamp,
          };
        })
        .filter(Boolean);

      const filteredHistory = formattedHistory.filter(
        (entry) =>
          entry.salesPersonName === salesPersonName &&
          entry.dateTime?.split(" ")[0].toString() ===
            formatDateDDMMYYYY(new Date())
      );

      const filteredHistoryData =
        userRole.toLowerCase() === "admin"
          ? formattedHistory
          : formattedHistory.filter(
              (entry) => entry.salesPersonName === salesPersonName
            );

      filteredHistory.sort((a, b) => {
        const parseGvizDate = (dateString) => {
          if (
            typeof dateString === "string" &&
            dateString.startsWith("Date(") &&
            dateString.endsWith(")")
          ) {
            const dateParts = dateString
              .substring(5, dateString.length - 1)
              .split(",");
            const year = parseInt(dateParts[0], 10);
            const month = parseInt(dateParts[1], 10);
            const day = parseInt(dateParts[2], 10);
            const hour = dateParts[3] ? parseInt(dateParts[3], 10) : 0;
            const minute = dateParts[4] ? parseInt(dateParts[4], 10) : 0;
            const second = dateParts[5] ? parseInt(dateParts[5], 10) : 0;
            return new Date(year, month, day, hour, minute, second);
          }
          return new Date(dateString);
        };
        const dateA = parseGvizDate(a._originalTimestamp);
        const dateB = parseGvizDate(b._originalTimestamp);
        return dateB.getTime() - dateA.getTime();
      });

      filteredHistoryData.sort((a, b) => {
        const parseGvizDate = (dateString) => {
          if (
            typeof dateString === "string" &&
            dateString.startsWith("Date(") &&
            dateString.endsWith(")")
          ) {
            const dateParts = dateString
              .substring(5, dateString.length - 1)
              .split(",");
            const year = parseInt(dateParts[0], 10);
            const month = parseInt(dateParts[1], 10);
            const day = parseInt(dateParts[2], 10);
            const hour = dateParts[3] ? parseInt(dateParts[3], 10) : 0;
            const minute = dateParts[4] ? parseInt(dateParts[4], 10) : 0;
            const second = dateParts[5] ? parseInt(dateParts[5], 10) : 0;
            return new Date(year, month, day, hour, minute, second);
          }
          return new Date(dateString);
        };
        const dateA = parseGvizDate(a._originalTimestamp);
        const dateB = parseGvizDate(b._originalTimestamp);
        return dateB.getTime() - dateA.getTime();
      });

      setAttendance(filteredHistory);
      setHistoryAttendance(filteredHistoryData);

      checkActiveSession(filteredHistory);
    } catch (error) {
      console.error("Error fetching attendance history:", error);
      showToast("Failed to load attendance history.", "error");
    } finally {
      setIsLoadingHistory(false);
    }
  };

  useEffect(() => {
    fetchAttendanceHistory();
  }, [currentUser, isAuthenticated]);

  const handleSubmit = async (e) => {
    e.preventDefault();

    if (!validateForm()) {
      showToast("Please fill in all required fields correctly.", "error");
      return;
    }

    if (!isAuthenticated || !currentUser || !salesPersonName) {
      showToast("User data not loaded. Please try logging in again.", "error");
      return;
    }

    if (formData?.status === "IN") {
      const indata = attendance.filter((item) => item.status === "IN");
      if (indata.length > 0) {
        showToast("Today Already in", "error");
        return;
      }
    }

    if (formData?.status === "OUT") {
      const indata = attendance.filter((item) => item.status === "IN");
      const outdata = attendance.filter((item) => item.status === "OUT");
      if (indata.length === 0) {
        showToast("First In", "error");
        return;
      }

      if (outdata.length > 0) {
        showToast("Today Already out", "error");
        return;
      }
    }

    setIsSubmitting(true);
    setIsGettingLocation(true);

    try {
      let currentLocation = null;
      try {
        currentLocation = await getCurrentLocation();
      } catch (locationError) {
        console.error("Location error:", locationError);
        showToast(locationError.message, "error");
        setIsSubmitting(false);
        setIsGettingLocation(false);
        return;
      }

      setIsGettingLocation(false);

      const currentDate = new Date();
      const timestamp = formatDateTime(currentDate);

      const dateForAttendance =
        formData.status === "IN" || formData.status === "OUT"
          ? formatDateTime(currentDate)
          : formData.startDate
          ? formatDateTime(new Date(formData.startDate + "T00:00:00"))
          : "";

      const endDateForLeave = formData.endDate
        ? formatDateTime(new Date(formData.endDate + "T00:00:00"))
        : "";

      let rowData = Array(10).fill("");
      rowData[0] = timestamp;
      rowData[1] = dateForAttendance;
      rowData[2] = endDateForLeave;
      rowData[3] = formData.status;
      rowData[4] = formData.reason;
      rowData[5] = currentLocation.latitude;
      rowData[6] = currentLocation.longitude;
      rowData[7] = currentLocation.mapLink;
      rowData[8] = currentLocation.formattedAddress;
      rowData[9] = salesPersonName;

      const payload = {
        sheetName: "Attendance",
        action: "insert",
        rowData: JSON.stringify(rowData),
      };

      const urlEncodedData = new URLSearchParams(payload);

      try {
        const response = await fetch(APPS_SCRIPT_URL, {
          method: "POST",
          headers: {
            "Content-Type": "application/x-www-form-urlencoded",
          },
          body: urlEncodedData,
        });

        console.log("response", response);

        const successMessage =
          formData.status === "IN"
            ? "Check-in successful!"
            : formData.status === "OUT"
            ? "Check-out successful!"
            : "Leave application submitted successfully!";
        showToast(successMessage, "success");

        setFormData({
          status: "",
          startDate: formatDateInput(new Date()),
          endDate: "",
          reason: "",
        });

        if (response.ok) {
          try {
            const responseText = await response.text();

            if (responseText.trim()) {
              const result = JSON.parse(responseText);
              if (result.success === false && result.activeSession) {
                await fetchAttendanceHistory();
                return;
              }
            }
          } catch (parseError) {
            console.log(
              "Response parsing issue, but success message already shown"
            );
          }
        }

        await fetchAttendanceHistory();
      } catch (fetchError) {
        console.error("Fetch error:", fetchError);

        const successMessage =
          formData.status === "IN"
            ? "Check-in successful!"
            : formData.status === "OUT"
            ? "Check-out successful!"
            : "Leave application submitted successfully!";
        showToast(successMessage, "success");

        setFormData({
          status: "",
          startDate: formatDateInput(new Date()),
          endDate: "",
          reason: "",
        });

        setTimeout(async () => {
          await fetchAttendanceHistory();
        }, 2000);
      }
    } catch (error) {
      console.error("Submission error:", error);
      showToast("Error recording attendance. Please try again.", "error");
    } finally {
      setIsSubmitting(false);
      setIsGettingLocation(false);
    }
  };

  const handleInputChange = (e) => {
    const { name, value } = e.target;
    
    if (name === "status" && value === "Leave") {
      if (hasCheckedInToday) {
        setFormData((prev) => ({
          ...prev,
          [name]: value,
          startDate: ""
        }));
      } else {
        setFormData((prev) => ({
          ...prev,
          [name]: value,
          startDate: formatDateInput(new Date())
        }));
      }
    } else {
      setFormData((prev) => ({
        ...prev,
        [name]: value,
      }));
    }

    if (errors[name]) {
      setErrors((prev) => ({
        ...prev,
        [name]: "",
      }));
    }
  };

  const showLeaveFields = formData.status === "Leave";

  if (!isAuthenticated || !currentUser || !currentUser.salesPersonName) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-slate-50 via-blue-50 to-indigo-50 flex items-center justify-center">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-red-600 mx-auto mb-4"></div>
          <p className="text-slate-600 font-medium">
            {!isAuthenticated
              ? "Please log in to view this page."
              : "Loading user data..."}
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-blue-50 to-indigo-50 p-0 lg:p-8">
      <div className="max-w-7xl mx-auto space-y-8">
        <div className="bg-white/80 backdrop-blur-sm rounded-2xl shadow-xl border border-white/20 overflow-hidden">
          <div className="bg-gradient-to-r from-blue-500 via-indigo-500 to-purple-500 px-8 py-6">
            <h3 className="text-2xl font-bold text-white mb-2">
              Mark Attendance
            </h3>
            <p className="text-emerald-50 text-lg">
              Record your daily attendance or apply for leave
            </p>
          </div>

          <form onSubmit={handleSubmit} className="space-y-8 p-8">
            <div className="grid gap-6 lg:grid-cols-1">
              <div className="space-y-2">
                <label className="block text-sm font-semibold text-slate-700 mb-3">
                  Status
                </label>
                <select
                  name="status"
                  value={formData.status}
                  onChange={handleInputChange}
                  className={`w-full px-4 py-3 bg-white border rounded-xl shadow-sm focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 transition-all duration-200 text-slate-700 font-medium ${
                    errors.status ? "border-red-300" : "border-slate-200"
                  }`}
                >
                  <option value="">Select status</option>
                  <option value="IN">IN</option>
                  <option value="OUT">OUT</option>
                  <option value="Leave">Leave</option>
                </select>
                {errors.status && (
                  <p className="text-red-500 text-sm mt-2 font-medium">
                    {errors.status}
                  </p>
                )}
              </div>
            </div>

            {!showLeaveFields && (
              <div className="bg-gradient-to-r from-emerald-50 to-teal-50 rounded-xl p-6 border border-emerald-100">
                <div className="text-sm font-semibold text-emerald-700 mb-2">
                  Current Date & Time
                </div>
                <div className="text-sm sm:text-2xl font-bold text-emerald-800">
                  {formatDateDisplay(new Date())}
                </div>
                {(formData.status === "IN" || formData.status === "OUT") && (
                  <div className="mt-3 text-sm text-emerald-600">
                    📍 Location will be automatically captured when you submit
                  </div>
                )}
              </div>
            )}

            {showLeaveFields && (
              <div className="bg-gradient-to-r from-amber-50 to-orange-50 rounded-xl p-0 sm:p-6 border border-amber-100 mb-6">
                <div className="text-sm font-semibold text-amber-700 mb-2">
                  Leave Application
                </div>
                <div className="text-lg font-bold text-amber-800">
                  {formatDateDisplay(new Date())}
                </div>
                <div className="mt-3 text-sm text-amber-600">
                  📍 Current location will be captured for leave application
                </div>
              </div>
            )}

            {showLeaveFields && (
              <div className="space-y-6">
                <div className="grid gap-6 lg:grid-cols-2">
                  <div className="space-y-2">
                    <label className="block text-sm font-semibold text-slate-700 mb-3">
                      Start Date
                    </label>
                    <input
                      type="date"
                      name="startDate"
                      value={formData.startDate}
                      onChange={handleInputChange}
                      className="w-full px-4 py-3 bg-white border border-slate-200 rounded-xl shadow-sm focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 transition-all duration-200 text-slate-700 font-medium"
                    />
                    {errors.startDate && (
                      <p className="text-red-500 text-sm mt-2 font-medium">
                        {errors.startDate}
                      </p>
                    )}
                  </div>

                  <div className="space-y-2">
                    <label className="block text-sm font-semibold text-slate-700 mb-3">
                      End Date
                    </label>
                    <input
                      type="date"
                      name="endDate"
                      value={formData.endDate}
                      onChange={handleInputChange}
                      min={formData.startDate}
                      className="w-full px-4 py-3 bg-white border border-slate-200 rounded-xl shadow-sm focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 transition-all duration-200 text-slate-700 font-medium"
                    />
                    {errors.endDate && (
                      <p className="text-red-500 text-sm mt-2 font-medium">
                        {errors.endDate}
                      </p>
                    )}
                  </div>
                </div>

                <div className="space-y-2">
                  <label className="block text-sm font-semibold text-slate-700 mb-3">
                    Reason
                  </label>
                  <textarea
                    name="reason"
                    value={formData.reason}
                    onChange={handleInputChange}
                    placeholder="Enter reason for leave"
                    className="w-full px-4 py-3 bg-white border border-slate-200 rounded-xl shadow-sm focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 transition-all duration-200 text-slate-700 font-medium min-h-32 resize-none"
                  />
                  {errors.reason && (
                    <p className="text-red-500 text-sm mt-2 font-medium">
                      {errors.reason}
                    </p>
                  )}
                </div>
              </div>
            )}

            <button
              type="submit"
              className="w-full lg:w-auto bg-gradient-to-r from-emerald-600 via-teal-600 to-cyan-600 hover:from-emerald-700 hover:via-teal-700 hover:to-cyan-700 text-white font-bold py-4 px-8 rounded-xl shadow-lg hover:shadow-xl transition-all duration-200 transform hover:scale-[1.02] disabled:opacity-50 disabled:cursor-not-allowed disabled:transform-none"
              disabled={
                isSubmitting ||
                isGettingLocation ||
                !currentUser?.salesPersonName
              }
            >
              {isGettingLocation ? (
                <span className="flex items-center gap-2">
                  <Loader2 className="h-5 w-5 animate-spin" />
                  Getting Location...
                </span>
              ) : isSubmitting ? (
                showLeaveFields ? (
                  "Submitting Leave..."
                ) : (
                  "Marking Attendance..."
                )
              ) : showLeaveFields ? (
                "Submit Leave Request"
              ) : (
                "Mark Attendance"
              )}
            </button>
          </form>
        </div>
      </div>

      <AttendanceHistory
        attendanceData={historyAttendance}
        isLoading={isLoadingHistory}
        userRole={userRole}
      />
    </div>
  );
};

export default Attendance;
