import { Injectable } from '@angular/core';
import { Client } from '@microsoft/microsoft-graph-client';
import { AuthService } from './auth.service';
import { AlertsService } from './alerts.service';
import  { Event, WorkbookSessionInfo } from '@microsoft/microsoft-graph-types'
import { DriveItem } from '@microsoft/microsoft-graph-types';

@Injectable({
  providedIn: 'root'
})
export class GraphService {

  private graphClient: Client;
  constructor(
    private authService: AuthService,
    private alertsService: AlertsService) {

    // Initialize the Graph client
    this.graphClient = Client.init({
      authProvider: async (done) => {
        // Get the token from the auth service
        let token = await this.authService.getAccessToken()
          .catch((reason) => {
            done(reason, null);
          });

        if (token)
        {
          done(null, token);
        } else {
          done("Could not get an access token", null);
        }
      }
    });
  }

  async getEvents(): Promise<Event[]> {
    try {
      let result =  await this.graphClient
        .api('/me/events')
        .select('subject,organizer,start,end')
        .orderby('createdDateTime DESC')
        .get();

      return result.value;
    } catch (error) {
      this.handleError('Could not get events', error);
    }
  }

  async getDriveChildren(): Promise<DriveItem[]> {
    try {
      let result = await this.graphClient
      .api('/me/drive/root/children')
      .get();
      return result.value;
    } catch(error) {
      this.handleError('Could not get files from drive', error);
    }
  }

  async getWorkbookSesson(id: string): Promise<WorkbookSessionInfo> {
    try {
      const workbookSessionInfo: WorkbookSessionInfo = {persistChanges: true};
      let result = await this.graphClient
      .api(`/me/drive/items/${id}/workbook/createSession`)
      .post(workbookSessionInfo);
      return result;
    } catch(error) {
      this.handleError('Unable to get workbook session', error);
    }
  }

  async getWorkbookSheets(id: string): Promise<any> {
    try {
      let result = await this.graphClient.api(`/me/drive/items/${id}/workbook/worksheets`).get();
      return result.value;
    } catch(error) {
      this.handleError('Unable to get workbook sheets', error);
    }
  }

  async updateWorkbookCell(id: string, sessionId: string, sheetName: string, cellAddress: string, cellValue: any) {
    try {
      let result = await this.graphClient
      .api(`/me/drive/items/${id}/workbook/worksheets/${sheetName}/range(address='${cellAddress}')`)
      .header('Workbook-Session-Id', sessionId)
      .patch({
        values: [[cellValue]]
      });
      return result;
    } catch(error) {
      this.handleError('unable to update cell', error);
    }
  }

  handleError(message: string, error: any) {
    this.alertsService.add(message, JSON.stringify(error, null, 2));
  }
}