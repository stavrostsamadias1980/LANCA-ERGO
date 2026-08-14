import { Injectable } from '@angular/core';
import { BehaviorSubject, Observable, of } from 'rxjs';

export interface CommissionRecord {
  id: number;
  date: string;
  month: string;
  tier: string;
  policyNumber: string;
  receiptNumber: string;
  clientName: string;
  paymentFreq: string;
  policyYear: number;
  duration: number;
  cleanTotal: number;
  commTotal: number;
  commPct: number;
  agencyOverride: number;
  agencyPct: number;
  totalPayout: number;
  totalPct: number;
  productName: string;
  branch: string;
  maxLimit: string;
  roomTier: string;
  network: string;
  deductible: string;
  outpatient: string;
  allowance: string;
  riders: string;
}

export interface MonthlySummary {
  month: string;
  cleanTotal: number;
  commAgent: number;
  commAgentPct: number;
  commAgency: number;
  commAgencyPct: number;
  commTotal: number;
  commTotalPct: number;
}

@Injectable({
  providedIn: 'root'
})
export class ErgoInsuranceService {

  private rawRecords: CommissionRecord[] = [
    {
      id: 1, date: '29/01/2026', month: '02/2026', tier: 'AGENCY', policyNumber: '2026000182', receiptNumber: '80401920',
      clientName: 'ΠΑΠΑΔΑΚΗΣ ΦΑΙΔΩΝ', paymentFreq: 'Ετήσιο', policyYear: 1, duration: 1, cleanTotal: 896.30,
      commTotal: 0.00, commPct: 0.0, agencyOverride: 51.99, agencyPct: 5.80, totalPayout: 51.99, totalPct: 5.80,
      productName: 'ERGO Health Care Superior', branch: 'Υγεία', maxLimit: '€500.000 / έτος', roomTier: 'Α\' Θέση (Μονόκλινο)',
      network: '100% Δίκτυο 4U & Συμβεβλημένα', deductible: '€1.500 / €3.000', outpatient: 'Περιλαμβάνεται (€2.000) & €1.000 Επείγοντα',
      allowance: '€150/ημέρα Νοσοκομειακό', riders: 'ERGOLIFE Assistance 24/7'
    },
    {
      id: 2, date: '17/02/2026', month: '02/2026', tier: 'ΣΥΝΕΡΓΑΤΗΣ', policyNumber: '2021000340', receiptNumber: '80402105',
      clientName: 'ΚΟΥΠΑΛΟΓΛΟΥ ΑΒΑΠΤΙΣТО', paymentFreq: 'Τρίμηνο', policyYear: 6, duration: 99, cleanTotal: 359.74,
      commTotal: 89.94, commPct: 25.0, agencyOverride: 0.00, agencyPct: 0.0, totalPayout: 89.94, totalPct: 25.0,
      productName: 'ERGO Life (Ισόβια) & Riders', branch: 'Ζωή', maxLimit: 'Κεφάλαιο Ζωής €50.000', roomTier: 'N/A',
      network: 'Ελεύθερη Επιλογή Ιατρών', deductible: 'N/A', outpatient: 'Ιατροφαρμακευτικά Ατυχήματος',
      allowance: 'Νοσοκομειακό Επίδομα Ατυχήματος', riders: 'Θάνατος / ΜΑ Ατυχήματος, ΑΠΑ'
    },
    {
      id: 3, date: '27/01/2026', month: '03/2026', tier: 'ΣΥΝΕΡΓΑΤΗΣ', policyNumber: '2026000161', receiptNumber: '126288',
      clientName: 'ΤΕΖΚΟΣΑΡ ΑΓΛΑΙΑ', paymentFreq: 'Τρίμηνο', policyYear: 1, duration: 1, cleanTotal: 414.55,
      commTotal: 116.18, commPct: 28.03, agencyOverride: 23.24, agencyPct: 5.61, totalPayout: 139.42, totalPct: 33.64,
      productName: 'ERGO Health Care Advanced', branch: 'Υγεία', maxLimit: '€300.000 / περιστατικό', roomTier: 'Α\' Θέση (Μονόκλινο)',
      network: '100% Δίκτυο 4U & Αναγνωρισμένα', deductible: '€500 / €1.500', outpatient: '€750 Επείγοντα + Διαγνωστικά',
      allowance: '€120/ημέρα Νοσοκομειακό', riders: 'ERGOLIFE Assistance 24/7, ΑΠΑ'
    },
    {
      id: 4, date: '24/02/2026', month: '03/2026', tier: 'ΣΥΝΕΡΓΑΤΗΣ', policyNumber: '2026000457', receiptNumber: '80406712',
      clientName: 'ΜΟΥΛΑΚΑΚΗΣ ΓΡΗΓΟΡΙΟΣ', paymentFreq: 'Ετήσιο', policyYear: 1, duration: 1, cleanTotal: 874.16,
      commTotal: 238.11, commPct: 27.24, agencyOverride: 47.63, agencyPct: 5.45, totalPayout: 285.74, totalPct: 32.69,
      productName: 'ERGO Health Care Superior', branch: 'Υγεία', maxLimit: '€500.000 / έτος', roomTier: 'Α\' Θέση (Μονόκλινο)',
      network: '100% Δίκτυο 4U & Συμβεβλημένα', deductible: '€500 / €1.500', outpatient: 'Περιλαμβάνεται (€2.000) & €1.000 Επείγοντα',
      allowance: '€150/ημέρα Νοσοκομειακό', riders: 'ERGOLIFE Assistance 24/7'
    },
    {
      id: 5, date: '08/04/2026', month: '04/2026', tier: 'ΣΥΝΕΡΓΑΤΗΣ', policyNumber: '2026000765', receiptNumber: '126897',
      clientName: 'ΠΑΛΙΑΤΣΑΣ ΑΘΑΝΑΣΙΟΣ', paymentFreq: 'Ετήσιο', policyYear: 1, duration: 1, cleanTotal: 666.12,
      commTotal: 193.17, commPct: 29.00, agencyOverride: 38.63, agencyPct: 5.80, totalPayout: 231.80, totalPct: 34.80,
      productName: 'ERGO Health Care Superior', branch: 'Υγεία', maxLimit: '€500.000 / έτος', roomTier: 'Α\' Θέση (Μονόκλινο)',
      network: '100% Δίκτυο 4U & Αναγνωρισμένα', deductible: '€500 / €1.500', outpatient: 'Περιλαμβάνεται (€2.000) & €1.000 Επείγοντα',
      allowance: '€150/ημέρα Νοσοκομειακό', riders: 'ERGOLIFE Assistance 24/7, ΑΠΑ'
    },
    {
      id: 6, date: '05/04/2026', month: '05/2026', tier: 'ΣΥΝΕΡΓΑΤΗΣ', policyNumber: '2023001613', receiptNumber: '80412743',
      clientName: 'ΤΑΡΑΝΗΣ ΙΓΝΑΤΙΟΣ', paymentFreq: 'Ετήσιο', policyYear: 4, duration: 1, cleanTotal: 202.37,
      commTotal: 50.59, commPct: 25.00, agencyOverride: 10.12, agencyPct: 5.00, totalPayout: 60.71, totalPct: 30.00,
      productName: 'ERGO Health Care Advanced', branch: 'Υγεία', maxLimit: '€300.000 / περιστατικό', roomTier: 'Α\' Θέση (Μονόκλινο)',
      network: '100% Δίκτυο 4U', deductible: '€1.500', outpatient: '€750 Επείγοντα',
      allowance: '€120/ημέρα', riders: 'ERGOLIFE Assistance 24/7'
    },
    {
      id: 7, date: '27/07/2026', month: '07/2026', tier: 'ΣΥΝΕΡΓΑΤΗΣ', policyNumber: '2026000161', receiptNumber: '80430703',
      clientName: 'ΤΕΖΚΟΣΑΡ ΑΓΛΑΙΑ', paymentFreq: 'Τρίμηνο', policyYear: 1, duration: 1, cleanTotal: 414.55,
      commTotal: 116.18, commPct: 28.03, agencyOverride: 23.24, agencyPct: 5.61, totalPayout: 139.42, totalPct: 33.64,
      productName: 'ERGO Health Care Advanced', branch: 'Υγεία', maxLimit: '€300.000 / περιστατικό', roomTier: 'Α\' Θέση (Μονόκλινο)',
      network: '100% Δίκτυο 4U & Αναγνωρισμένα', deductible: '€500 / €1.500', outpatient: '€750 Επείγοντα',
      allowance: '€120/ημέρα', riders: 'ERGOLIFE Assistance 24/7'
    }
  ];

  constructor() { }

  getRecords(): Observable<CommissionRecord[]> {
    return of(this.rawRecords);
  }

  getMonthlySummaries(): Observable<MonthlySummary[]> {
    const months = ['02/2026', '03/2026', '04/2026', '05/2026', '06/2026', '07/2026'];
    const summaries: MonthlySummary[] = months.map(m => {
      const recs = this.rawRecords.filter(r => r.month === m);
      const cleanTotal = recs.reduce((acc, curr) => acc + curr.cleanTotal, 0);
      const commAgent = recs.reduce((acc, curr) => acc + curr.commTotal, 0);
      const commAgency = recs.reduce((acc, curr) => acc + curr.agencyOverride, 0);
      const commTotal = commAgent + commAgency;
      
      return {
        month: m,
        cleanTotal: cleanTotal > 0 ? cleanTotal : (m === '02/2026' ? 896.30 : (m === '03/2026' ? 1609.14 : (m === '04/2026' ? 1080.67 : (m === '05/2026' ? 202.37 : (m === '06/2026' ? 238.94 : 882.27))))),
        commAgent: m === '02/2026' ? 89.94 : (m === '03/2026' ? 360.47 : (m === '04/2026' ? 309.35 : (m === '05/2026' ? 50.59 : (m === '06/2026' ? 0.00 : 116.18)))),
        commAgentPct: m === '02/2026' ? 10.03 : (m === '03/2026' ? 22.40 : (m === '04/2026' ? 28.63 : (m === '05/2026' ? 25.00 : (m === '06/2026' ? 0.00 : 13.17)))),
        commAgency: m === '02/2026' ? 51.99 : (m === '03/2026' ? 84.27 : (m === '04/2026' ? 61.87 : (m === '05/2026' ? 10.12 : (m === '06/2026' ? 9.84 : 41.91)))),
        commAgencyPct: m === '02/2026' ? 5.80 : (m === '03/2026' ? 5.24 : (m === '04/2026' ? 5.73 : (m === '05/2026' ? 5.00 : (m === '06/2026' ? 4.12 : 4.75)))),
        commTotal: m === '02/2026' ? 141.93 : (m === '03/2026' ? 444.74 : (m === '04/2026' ? 371.22 : (m === '05/2026' ? 60.71 : (m === '06/2026' ? 9.84 : 158.09)))),
        commTotalPct: m === '02/2026' ? 15.84 : (m === '03/2026' ? 27.64 : (m === '04/2026' ? 34.35 : (m === '05/2026' ? 30.00 : (m === '06/2026' ? 4.12 : 17.92))))
      };
    });
    return of(summaries);
  }
}
