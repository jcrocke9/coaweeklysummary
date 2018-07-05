import * as React from 'react';
import styles from './CoaWeeklySummary.module.scss';
import { ICoaWeeklySummaryProps } from './ICoaWeeklySummaryProps';
import { ICoaWeeklySummaryState } from './ICoaWeeklySummaryState';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IListItem } from './IListItem';
import { HealthIndicators } from './HealthIndicators';

export default class CoaWeeklySummary extends React.Component<ICoaWeeklySummaryProps, ICoaWeeklySummaryState> {
  constructor(props: ICoaWeeklySummaryProps, state: ICoaWeeklySummaryState) {
    super(props);
    this.onChange_reportPeriodEnd = this.onChange_reportPeriodEnd.bind(this);
    this.state = {
      reportPeriodEnd: '',
      Report_x0020_Period_x0020_End: '',
      weeklySummary: ''
    };
  }

  public componentDidMount(): void {
    this.readItem();
  }
  public onChange_reportPeriodEnd(Report_x0020_Period_x0020_End: string): void {
    this.setState({
      reportPeriodEnd: this.dateToDate(Report_x0020_Period_x0020_End)
    });
  }
  public render(): React.ReactElement<ICoaWeeklySummaryProps> {
    const reportPeriodEnd: string = this.state.reportPeriodEnd;
    const weeklySummary: string = this.state.weeklySummary;
    return (
      <div className={styles.coaWeeklySummary}>
        <div className="ms-Grid">
        <div className={styles.row}>
        <div className={styles.column}>
        <span className={styles.title}>Weekly Status for Report Period Ending: {reportPeriodEnd}</span>
        </div>
        </div>
        <div className={styles.row}>
          <div className={styles.column}>

            <p className={styles.description} dangerouslySetInnerHTML={{ __html: weeklySummary }}></p>
          </div>
          <div className={styles.leftColumn}>
            <HealthIndicators siteUrl={this.props.siteUrl} spHttpClient={this.props.spHttpClient} spSiteUrl={this.props.spSiteUrl} />
          </div>
        </div>
        </div>
      </div>
    );
  }

  private readItem(): void {
    this.getLatestItemId()
      .then((itemId: number): Promise<SPHttpClientResponse> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }
        return this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('Weekly Status')/items(${itemId})?$select=Title,Id,Report_x0020_Period_x0020_End,Weekly_x0020_Summary`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        return response.json();
      })
      .then((item: IListItem): void => {
        this.setState({
          Report_x0020_Period_x0020_End: item.Report_x0020_Period_x0020_End,
          weeklySummary: item.Weekly_x0020_Summary
        });
        this.onChange_reportPeriodEnd(item.Report_x0020_Period_x0020_End);
      }, (error: any): void => {
        this.setState({
          weeklySummary: '<p>No weekly status found</p>'
        });
      });
  }
  private getLatestItemId(): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('Weekly Status')/items?$orderby=Id desc&$top=1&$select=id`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: { value: { Id: number }[] }): void => {
          if (response.value.length === 0) {
            resolve(-1);
          }
          else {
            resolve(response.value[0].Id);
          }
        });
    });
  }
  public dateToDate(strDate: string): string {
    let dateValue: Date = new Date(strDate);
    let dateValueFormatted: string = dateValue.toDateString();
    return dateValueFormatted;
  }
}
