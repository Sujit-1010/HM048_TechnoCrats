import React, { useState, useEffect, useCallback } from "react";
import { useLocation, useNavigate } from "react-router-dom";
import { Session, Batch } from "@/entities/all";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { ArrowLeft, RefreshCw, Clock } from "lucide-react";
import { createPageUrl } from "@/utils";
import { Skeleton } from "@/components/ui/skeleton";

const QR_EXPIRY_SECONDS = 45;

export default function QRAttendancePage() {
  const location = useLocation();
  const navigate = useNavigate();
  const [session, setSession] = useState(null);
  const [batch, setBatch] = useState(null);
  const [qrCodeDataUrl, setQrCodeDataUrl] = useState("");
  const [countdown, setCountdown] = useState(QR_EXPIRY_SECONDS);
  const [isLoading, setIsLoading] = useState(true);

  // Get session_id from query parameters
  const urlParams = new URLSearchParams(location.search);
  const session_id = urlParams.get('session_id');

  const generateQRCodeData = useCallback(() => {
    if (!session || !batch) return;
    
    const timestamp = Date.now();
    const url = new URL(window.location.origin + createPageUrl("QRAttendanceStudent"));
    url.searchParams.append("session_id", session.id);
    url.searchParams.append("timestamp", timestamp);
    
    if (batch.batch_type === 'private') {
      url.searchParams.append("batch_type", "private");
    }
    
    // Use an external API to generate the QR code image
    const qrApiUrl = `https://api.qrserver.com/v1/create-qr-code/?data=${encodeURIComponent(url.toString())}&size=256x256&bgcolor=ffffff`;
    
    setQrCodeDataUrl(qrApiUrl);
    setCountdown(QR_EXPIRY_SECONDS);
  }, [session, batch]);

  useEffect(() => {
    const loadSessionAndBatch = async () => {
      if (!session_id) {
        console.error("No session_id found in URL parameters");
        setIsLoading(false);
        return;
      }

      setIsLoading(true);
      try {
        console.log("Loading session with ID:", session_id);
        const sessionData = await Session.get(session_id);
        if (sessionData && sessionData.batch_id) {
          setSession(sessionData);
          const batchData = await Batch.filter({ batch_id: sessionData.batch_id });
          if (batchData.length > 0) {
            setBatch(batchData[0]);
          } else {
            setBatch({ batch_type: 'standard' }); // Fallback
          }
        }
      } catch (error) {
        console.error("Error loading session/batch data:", error);
      }
      setIsLoading(false);
    };
    loadSessionAndBatch();
  }, [session_id]);

  useEffect(() => {
    if (session && batch) {
      generateQRCodeData();
      const timer = setInterval(generateQRCodeData, QR_EXPIRY_SECONDS * 1000);
      return () => clearInterval(timer);
    }
  }, [session, batch, generateQRCodeData]);

  useEffect(() => {
    const countdownTimer = setInterval(() => {
      setCountdown(prev => (prev > 0 ? prev - 1 : 0));
    }, 1000);
    return () => clearInterval(countdownTimer);
  }, [qrCodeDataUrl]);

  if (!session_id) {
    return (
      <div className="p-8 text-center">
        <h2 className="text-xl font-semibold text-red-600">Error: No Session ID Provided</h2>
        <p className="text-gray-600 mt-2">Please navigate back and try again.</p>
        <Button onClick={() => navigate(createPageUrl("Attendance"))} className="mt-4">
          Back to Attendance
        </Button>
      </div>
    );
  }

  if (isLoading) {
    return (
        <div className="p-8"><Skeleton className="h-96 w-full" /></div>
    );
  }

  if (!session) {
    return <div className="p-8 text-center text-red-500">Session not found.</div>;
  }

  return (
    <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4">
      <Card className="w-full max-w-md shadow-2xl">
        <CardHeader className="text-center">
          <Button variant="ghost" size="icon" className="absolute top-4 left-4" onClick={() => navigate(createPageUrl("Attendance"))}>
            <ArrowLeft />
          </Button>
          <CardTitle className="text-2xl font-bold">Scan for Attendance</CardTitle>
          <p className="text-slate-600">{session.session_name}</p>
        </CardHeader>
        <CardContent className="flex flex-col items-center gap-6">
          <div className="bg-white p-4 rounded-lg border-4 border-slate-800">
            {qrCodeDataUrl ? (
              <img src={qrCodeDataUrl} alt="Attendance QR Code" width={256} height={256} />
            ) : (
              <Skeleton className="w-64 h-64" />
            )}
          </div>
          <div className="text-center">
            <p className="flex items-center gap-2 text-lg font-semibold text-red-600">
              <Clock className="w-5 h-5 animate-spin" />
              QR code refreshes in: {countdown}s
            </p>
            <p className="text-xs text-slate-500 mt-1">Point your camera at the QR Code to mark your attendance.</p>
          </div>
          <Button variant="outline" onClick={generateQRCodeData}>
            <RefreshCw className="w-4 h-4 mr-2" />
            Refresh QR Manually
          </Button>
        </CardContent>
      </Card>
    </div>
  );
}