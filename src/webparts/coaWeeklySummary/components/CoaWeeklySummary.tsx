import * as React from 'react';
import styles from './CoaWeeklySummary.module.scss';
import { ICoaWeeklySummaryProps } from './ICoaWeeklySummaryProps';
import { ICoaWeeklySummaryState } from './ICoaWeeklySummaryState';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IListItem } from './IListItem';

export default class CoaWeeklySummary extends React.Component<ICoaWeeklySummaryProps, ICoaWeeklySummaryState> {
  constructor(props: ICoaWeeklySummaryProps, state: ICoaWeeklySummaryState) {
    super(props);

    this.state = {
      items: []
    };
  }
  public render(): React.ReactElement<ICoaWeeklySummaryProps> {
    const items: JSX.Element[] = this.state.items.map((item: IListItem, i: number): JSX.Element => {
      return (
        <div className={ styles.column }>
        <span className={ styles.title }>Weekly Status for Report Period Ending: {escape(item.reportPeriodEnd)}</span>
        <p className={ styles.description }>{escape(item.weeklySummary)}</p>
      </div>
      );
    });
    return (
      <div className={ styles.coaWeeklySummary }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            {items}
          </div>
        </div>
      </div>
    );
  }

  private readItem(): void {
    this.setState({
      items: []
    });
    this.getLatestItemId()
      .then((itemId: number): Promise<SPHttpClientResponse> => {
        if (itemId === -1) {
          throw new Error('No items found in the list');
        }

        this.setState({
          items: []
        });
        return this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items(${itemId})?$select=Title,Id`,
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
          items: []
        });
      }, (error: any): void => {
        this.setState({
          items: []
        });
      });
  }
  private getLatestItemId(): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items?$orderby=Id desc&$top=1&$select=id`,
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
}
