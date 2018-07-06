import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './CoaWeeklySummary.module.scss';
export interface IHealthIndicatorsProps {
    spHttpClient?: any;
    siteUrl?: string;
    spSiteUrl?: string;
}
export interface IHealthIndicatorsState {
    projectUid?: string;
    overallProjectHealth?: string;
    scopeHealth?: string;
    resourceHealth?: string;
    technicalHealth?: string;
}
export interface IProjectProperties {
    MSPWAPROJUID: string;
}
export interface IHealthIndicator {
    Custom_x005f_1ae9c5dbc17ce41180d9005056b25256?: any[]; // overallProjectHealth
    Custom_x005f_1f24ced7c27ce41180d9005056b25256?: any[]; // scopeHealth
    Custom_x005f_50cb138cc27ce41180d9005056b25256?: any[]; // resourceHealth
    Custom_x005f_9fe93834c57ce41180d9005056b25256?: any[]; // technicalHealth
}
export class HealthIndicators extends React.Component<IHealthIndicatorsProps, IHealthIndicatorsState> {

    constructor(props: IHealthIndicatorsProps, state: IHealthIndicatorsState) {
        super(props);
        this.onChange_projectUid = this.onChange_projectUid.bind(this);
        this.onChange_overallProjectHealth = this.onChange_overallProjectHealth.bind(this);
        this.onChange_scopeHealth = this.onChange_scopeHealth.bind(this);
        this.onChange_resourceHealth = this.onChange_resourceHealth.bind(this);
        this.onChange_technicalHealth = this.onChange_technicalHealth.bind(this);
        this.state = {
            projectUid: '',
            overallProjectHealth: '',
            scopeHealth: '',
            resourceHealth: '',
            technicalHealth: ''
        };
    }
    public componentDidMount(): void {
        this.getProjectUid();
    }
    public onChange_projectUid(projectUid: string): void {
        this.setState({
            projectUid
        });
        this.getOverallProjectHealthIndicator();
        this.getScopeHealthIndicator();
        this.getResourceHealthIndicator();
        this.getTechnicalHealthIndicator();
    }
    public onChange_overallProjectHealth(overallProjectHealth: string): void {
        this.setState({overallProjectHealth});
    }
    public onChange_scopeHealth(scopeHealth: string): void {
        this.setState({scopeHealth});
    }
    public onChange_resourceHealth(resourceHealth: string): void {
        this.setState({resourceHealth});
    }
    public onChange_technicalHealth(technicalHealth: string): void {
        this.setState({technicalHealth});
    }
    public render() {
        const overallProjectHealthState: string = this.state.overallProjectHealth;
        let overallProjectHealth: string = this.determineHealthIndicator(overallProjectHealthState);
        const scopeHealthState: string = this.state.scopeHealth;
        let scopeHealth: string = this.determineHealthIndicator(scopeHealthState);
        const resouceHealthState: string = this.state.resourceHealth;
        let resouceHealth: string = this.determineHealthIndicator(resouceHealthState);
        const technicalHealthState: string = this.state.technicalHealth;
        let technicalHealth: string = this.determineHealthIndicator(technicalHealthState);
        return (
            <div className={styles.rightJust}>
                <div className={styles.healthIndicators}>Overall Project Health <i className={overallProjectHealth}></i></div>
                <div className={styles.healthIndicators}>Scope Health <i className={scopeHealth}></i></div>
                <div className={styles.healthIndicators}>Resource Health <i className={resouceHealth}></i></div>
                <div className={styles.healthIndicators}>Technical Health <i className={technicalHealth}></i></div>
            </div>
        );
    }
    private getProjectUid(): Promise<String> {
        return new Promise<String>((resolve: (itemId: String) => void, reject: (error: any) => void): void => {
            this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/AllProperties?$select=MSPWAPROJUID`,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'odata-version': ''
                    }
                })
                .then((response: SPHttpClientResponse): Promise<{ value: { Id: String }[] }> => {
                    return response.json();
                }, (error: any): void => {
                    reject(error);
                })
                .then((item: IProjectProperties): void => {
                    this.setState({
                        projectUid: item.MSPWAPROJUID
                    });
                    this.onChange_projectUid(item.MSPWAPROJUID);
                });
        });
    }
    private getOverallProjectHealthIndicator(): Promise<String> {
        return this.props.spHttpClient.get(`${this.props.spSiteUrl}/_api/ProjectServer/Projects('${this.state.projectUid}')/IncludeCustomFields?$select=Custom_x005f_1ae9c5dbc17ce41180d9005056b25256`,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            })
            .then((response: SPHttpClientResponse): Promise<IHealthIndicator> => {
                return response.json();
            })
            .then((item: IHealthIndicator): void => {
                this.setState({
                    overallProjectHealth: item.Custom_x005f_1ae9c5dbc17ce41180d9005056b25256[0]
                });
                this.onChange_overallProjectHealth(item.Custom_x005f_1ae9c5dbc17ce41180d9005056b25256[0]);
            });
    }
    private getScopeHealthIndicator(): Promise<String> {
        return this.props.spHttpClient.get(`${this.props.spSiteUrl}/_api/ProjectServer/Projects('${this.state.projectUid}')/IncludeCustomFields?$select=Custom_x005f_1f24ced7c27ce41180d9005056b25256`,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            })
            .then((response: SPHttpClientResponse): Promise<IHealthIndicator> => {
                return response.json();
            })
            .then((item: IHealthIndicator): void => {
                this.setState({
                    scopeHealth: item.Custom_x005f_1f24ced7c27ce41180d9005056b25256[0]
                });
                this.onChange_scopeHealth(item.Custom_x005f_1f24ced7c27ce41180d9005056b25256[0]);
            });
    }
    private getResourceHealthIndicator(): Promise<String> {
        return this.props.spHttpClient.get(`${this.props.spSiteUrl}/_api/ProjectServer/Projects('${this.state.projectUid}')/IncludeCustomFields?$select=Custom_x005f_50cb138cc27ce41180d9005056b25256`,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            })
            .then((response: SPHttpClientResponse): Promise<IHealthIndicator> => {
                return response.json();
            })
            .then((item: IHealthIndicator): void => {
                this.setState({
                    resourceHealth: item.Custom_x005f_50cb138cc27ce41180d9005056b25256[0]
                });
                this.onChange_resourceHealth(item.Custom_x005f_50cb138cc27ce41180d9005056b25256[0]);
            });
    }
    private getTechnicalHealthIndicator(): Promise<String> {
        return this.props.spHttpClient.get(`${this.props.spSiteUrl}/_api/ProjectServer/Projects('${this.state.projectUid}')/IncludeCustomFields?$select=Custom_x005f_9fe93834c57ce41180d9005056b25256`,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            })
            .then((response: SPHttpClientResponse): Promise<IHealthIndicator> => {
                return response.json();
            })
            .then((item: IHealthIndicator): void => {
                this.setState({
                    technicalHealth: item.Custom_x005f_9fe93834c57ce41180d9005056b25256[0]
                });
                this.onChange_technicalHealth(item.Custom_x005f_9fe93834c57ce41180d9005056b25256[0]);
            });
    }
    private determineHealthIndicator(entry: string): string {
        var toReturn: string;
        switch (entry) {
            case 'Entry_4c5954c9a77ce41180d9005056b25256':
                toReturn = `ms-Icon ms-Icon--SkypeCircleCheck ${styles.green}`;
                break;
            case 'Entry_4b5954c9a77ce41180d9005056b25256':
                toReturn = `ms-Icon ms-Icon--SkypeCircleClock ${styles.amber}`;
                break;
            case 'Entry_4a5954c9a77ce41180d9005056b25256':
                toReturn = `ms-Icon ms-Icon--SkypeCircleMinus ${styles.red}`;
                break;
            default:
                toReturn = `ms-Icon ms-Icon--SkypeCircleClock ${styles.grey}`;
                break;
        }
        return toReturn;
    }
}
