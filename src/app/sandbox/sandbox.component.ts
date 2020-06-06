import { DriveItem } from './../drive-item';
import { AlertsService } from './../alerts.service';
import { GraphService } from './../graph.service';
import { Component, OnInit } from '@angular/core';
import { FormGroup, FormBuilder } from '@angular/forms';

@Component({
  selector: 'app-sandbox',
  templateUrl: './sandbox.component.html',
  styleUrls: ['./sandbox.component.css']
})
export class SandboxComponent implements OnInit {
  drive: DriveItem[] = [];
  form: FormGroup;
  constructor(private graphService: GraphService,
    private alertsService: AlertsService, private fb: FormBuilder) { }

  ngOnInit(): void {
    this.graphService.getDriveChildren().then(drive => {
      this.drive = drive;
    });

    this.form = this.fb.group({
      'id': [],
      'sheetName': [],
      'cellAddress': [],
      'cellValue': []
    });
  }

  async formSubmit() {
    console.log(this.form.value);
    const fileId = this.form.value.id;
    const sheetName = this.form.value.sheetName;
    const cellAddress = this.form.value.cellAddress;
    const cellValue = this.form.value.cellValue;
    const session = await this.graphService.getWorkbookSesson(fileId);
    //const sheets = await this.graphService.getWorkbookSheets(this.form.value.id);
    const update = await this.graphService.updateWorkbookCell(fileId, session.id, sheetName, cellAddress, cellValue);
    console.log(update);
  }

}
