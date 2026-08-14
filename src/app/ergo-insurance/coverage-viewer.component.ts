import { Component, OnInit } from '@angular/core';
import { ErgoInsuranceService, CommissionRecord } from './ergo-insurance.service';

@Component({
  selector: 'app-coverage-viewer',
  templateUrl: './coverage-viewer.component.html',
  styleUrls: ['./coverage-viewer.component.css']
})
export class CoverageViewerComponent implements OnInit {

  records: CommissionRecord[] = [];
  selectedRecord: CommissionRecord | null = null;

  constructor(private ergoService: ErgoInsuranceService) { }

  ngOnInit(): void {
    this.ergoService.getRecords().subscribe(data => {
      this.records = data;
      if (data.length > 0) {
        this.selectedRecord = data[0];
      }
    });
  }

  selectPolicy(rec: CommissionRecord) {
    this.selectedRecord = rec;
  }
}
