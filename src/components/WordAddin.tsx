import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { DocumentCard, DocumentCardTitle } from 'office-ui-fabric-react/lib/DocumentCard';

import * as OfficeHelpers from '@microsoft/office-js-helpers';
 
import { IOfficeAddinProps } from './OfficeAddin';
import { access } from 'fs';

export interface IWordAddinProps extends IOfficeAddinProps {

}

interface IWordAddInState {
    data: any;
    mode: any;
}

export default class WordAddin extends React.Component<IWordAddinProps, IWordAddInState> {

    private provider: string = OfficeHelpers.DefaultEndpoints.AzureAD;;
    constructor(props: any) {
        super(props);
        this.state = {
            data: '',
            mode: 'LOADING_DATA'
        }
    }

    componentDidMount(): void {
        this.loadContent();
    }

    loadContent = () => {
        let self = this;

        let resourceUrl:string = "https://pjsummersjr2.ngrok.io/api/opportunities";
        self.props.authenticator.authenticate(this.provider, false).then(
            function(response: any){
                console.debug(`Requesting data from ${resourceUrl}`);
                let accessToken = response.access_token;
                let requestConfig = {
                    method:"GET",
                    headers: {
                        "Authorization": `Bearer ${accessToken}`,
                        "Content-Type": `application/json`                    }
                }
                fetch(resourceUrl, requestConfig)
                .then(response => response.json()) //Not quite sure why the multiple steps are required
                .then(
                    function(response){
                        //console.log(JSON.stringify(response));
                        self.setState({
                            data:response
                        });
                    },
                    (error) => {
                        console.log(`Error from authenticate: ${error}`);
                        var token = self.props.authenticator.tokens.get(self.provider);
                        console.log(`Got a token: ${JSON.stringify(token)}`);
                    }
                )
            }
        )
        .catch(function(error:any){
            console.error(`Error caught in code: ${error}`);
        });
    }

    drawOpportunities(oppData: any): any {
        let oppContent = (<div>No opportunity data available</div>);
        if(!oppData) return oppContent;
        oppContent = oppData.value.map((item: any, index: number) =>{
            return (<DocumentCard key={item.opportunityid}>
                        <DocumentCardTitle title={item.name} shouldTruncate={false} />
                        <DocumentCardTitle title={item.description ? item.description : 'No description found'} shouldTruncate={true} showAsSecondaryTitle={true} />
                    </DocumentCard>)
        })
        return (oppContent);
    }

    render() {
        if(!this.props.isOfficeInitialized) {
            return (<ProgressIndicator label="No Office environment detected" description="Please load this page within an Office add-in"/>);
        }
        let opportunities: any = this.drawOpportunities(this.state.data);
        return (<div>{opportunities}</div>);
    }

}