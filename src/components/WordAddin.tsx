import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { DocumentCard, DocumentCardTitle, DocumentCardActions } from 'office-ui-fabric-react/lib/DocumentCard';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';


import * as OfficeHelpers from '@microsoft/office-js-helpers';
 
import { IOfficeAddinProps, OfficeAddInComponent, IOfficeAddinState } from './OfficeAddin';

import * as q from 'q';

export interface IWordAddinProps extends IOfficeAddinProps {

}
/**
 * oppData - contains the JSON from the API endpoint
 * mode - identifies the state of the add-in
 * status - a string identifying the status of operations. Mainly for error or 'loading' states.
 */
interface IWordAddInState extends IOfficeAddinState{
    oppData: any;
    mode: string;
    status: string;
}
/**
 * LOADING_OPPS - opportunity list data is being retrieved from API
 * LOADING_OPPDETAIL - data for a single opportunity is being retrieved from the API
 * OPP_LIST - opportunity list is being displayed
 * OPP_DETAIL - single opportunity detail data is being displayed
 */
enum OPP_MODE {
    LOADING_OPPS = "loading-opps",
    LOADING_OPPDETAIL = "loading-opp-detail",
    OPP_LIST = "opp-list",
    OPP_DETAIL = "opp-detail",
    ERROR = "error"
}

export default class WordAddin<IWordAddinProps, IWordAddInState> extends OfficeAddInComponent {    

    constructor(props: IOfficeAddinProps, state: IOfficeAddinState) {
        super(props, state);
        this.provider = OfficeHelpers.DefaultEndpoints.AzureAD;
        this.state = {
            oppData: '',
            mode: OPP_MODE.LOADING_OPPS,
            status:'Loading opportunities'
        }

        initializeIcons();
    }
    /**
     * When the component loads, it will load whatever loads from this method
     * 
     * This is a base React component method so I don't want to put too much application-specific logic/code in here
     */
    componentDidMount(): void {
        this.loadOpportunities();
    }
    /**
     * Loads a list of opportunities into the state object - this will call render, FYI
     */
    loadOpportunities(): void {
        let opportunityUrl = "https://pjsummersjr2.ngrok.io/api/opportunities"
        let self = this;
        this.loadContent(opportunityUrl)
        .then(
            function(response: any){
                self.setState({
                    oppData:response,
                    mode: OPP_MODE.OPP_LIST
                });
            },
            (error: any) => {
                console.log(`Error from authenticate: ${error}`);
                var token = self.props.authenticator.tokens.get(self.provider);
                console.log(`Got a token: ${JSON.stringify(token)}`);
            } 
        );
    }

    /**
     * Renders the JSX for the opportunity list
     * @param oppData - JSON data representing a list of opportunities from the API endpoint
     */
    renderOpportunities(oppData: any): any {
        let oppContent = (<div>No opportunity data available</div>);
        if(!oppData) return oppContent;
        oppContent = oppData.value.map((item: any, index: number) =>{
            return (
                <DocumentCard key={item.opportunityid}>
                    <DocumentCardTitle title={item.name} shouldTruncate={false} />
                    <DocumentCardTitle title={item.description ? item.description : 'No description found'} shouldTruncate={true} showAsSecondaryTitle={true} />
                    <DocumentCardActions 
                        actions={[              
                            {
                                iconProps: {iconName: 'Dynamics365Logo'},
                                text: 'Open in Dynamics 365',
                                onClick: (ev: any) => {
                                window.open('https://paulsumm.crm.dynamics.com')
                                },
                                ariaLabel: 'Open in Dynamics 365'
                            },
                            {
                                iconProps: {iconName: 'CirclePlus'},
                                text: 'Show Opportunity Details',
                                onClick: (ev: any) => {
                                this.loadOpportunityDetails(item.opportunityid);
                                },
                                ariaLabel: 'Click for opportunity details'
                            }
                        ]}
                    />
                </DocumentCard>
            )
        })
        return oppContent;
    }

    loadOpportunityDetails(oppId: string): void {
        let opportunityUrl = `https://pjsummersjr2.ngrok.io/api/opportunities/${oppId}`
        let self = this;
        this.loadContent(opportunityUrl)
        .then(
            function(response: any){
                self.setState({
                    oppData:response,
                    mode: OPP_MODE.OPP_DETAIL
                });
            },
            (error: any) => {
                console.log(`Error from authenticate: ${error}`);
                var token = self.props.authenticator.tokens.get(self.provider);
                console.log(`Got a token: ${JSON.stringify(token)}`);
            } 
        );
    }

    renderOpportunityDetail(oppData: any): any {
        return (<div>{JSON.stringify(oppData)}</div>);
    }

    render() {
        if(!this.props.isOfficeInitialized) {
            return (<ProgressIndicator label="No Office environment detected" description="Please load this page within an Office add-in"/>);
        }
        if(this.state.mode === OPP_MODE.LOADING_OPPS) {
            return (<ProgressIndicator label={this.state.status}/>);
        }
        let content: any = (<div>Something weird happened. No data available</div>) 
        if(this.state.mode == OPP_MODE.OPP_LIST) content = this.renderOpportunities(this.state.oppData);
        if(this.state.mode == OPP_MODE.OPP_DETAIL) content = this.renderOpportunityDetail(this.state.oppData);
        return (<div>{content}</div>);
    }

}