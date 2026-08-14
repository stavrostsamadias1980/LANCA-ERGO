import { Component, OnInit, ViewChild } from '@angular/core';
import { MatTableDataSource } from '@angular/material/table';
import { MatPaginator } from '@angular/material/paginator';
import { MatSort } from '@angular/material/sort';
import { ErgoInsuranceService, CommissionRecord } from './ergo-insurance.service';

@Component({
  selector: 'app-commissions-table',
  templateUrl: './commissions-table.component.html',
  styleUrls: ['./commissions-table.component.css']
})
export class CommissionsTableComponent implements OnInit {

  displayedColumns: string[] = [
    'date', 'month', 'policyNumber', 'clientName', 'productName',
    'cleanTotal', 'commTotal', 'commPct', 'agencyOverride', 'agencyPct', 'totalPayout'
  ];

  dataSource: MatTableDataSource<CommissionRecord> = new MatTableDataSource<CommissionRecord>([]);
  selectedMonth: string = 'ALL';
  selectedTier: string = 'ALL';

  @ViewChild(MatPaginator) paginator!: MatPaginator;
  @ViewChild(MatSort) sort!: MatSort;

  constructor(private ergoService: ErgoInsuranceService) { }

  ngOnInit(): void {
    this.ergoService.getRecords().subscribe(data => {
      this.dataSource.data = data;
    });
  }

  ngAfterViewInit() {
    this.dataSource.paginator = this.paginator;
    this.dataSource.sort = this.sort;
  }

  applyFilter(event: Event) {
    const filterValue = (event.target as HTMLInputElement).value;
    this.dataSource.filter = filterValue.trim().toLowerCase();
  }

  filterByMonth(month: string) {
    this.selectedMonth = month;
    if (month === 'ALL') {
      this.ergoService.getRecords().subscribe(data => this.dataSource.data = data);
    } else {
      this.ergoService.getRecords().subscribe(data => {
        this.dataSource.data = data.filter(r => r.month === month);
      });
    }
  }
}
