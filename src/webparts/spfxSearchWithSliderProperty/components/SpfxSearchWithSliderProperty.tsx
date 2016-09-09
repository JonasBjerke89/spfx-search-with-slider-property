require('set-webpack-public-path!');

import * as React from 'react';
import {
  Spinner,
  DocumentCard,
  DocumentCardPreview,
  DocumentCardTitle,
  DocumentCardActivity,
 } from 'office-ui-fabric-react';

import {
  IWebPartContext
} from '@microsoft/sp-client-preview';

import { EnvironmentType } from '@microsoft/sp-client-base';

import MockHttpClient from './../Utilities/MockHttpClient';

import styles from '../SpfxSearchWithSliderProperty.module.scss';
import { ISpfxSearchWithSliderPropertyWebPartProps } from '../ISpfxSearchWithSliderPropertyWebPartProps';

export interface ISpfxSearchWithSliderPropertyProps extends ISpfxSearchWithSliderPropertyWebPartProps {
  siteUrl: string;
  context: IWebPartContext;
}

/* Interface to define our React state object. Add this to the React.Component method as a second paramter to handle state */
export interface ISearchResultsState {
  searchResults: ISearchResult[];
  loading: boolean;
}

export interface ISearchResult {
  title: string;
  url: string;
  key: string;
  previewImageUrl: string;
  extension: string;
}

export interface IResult {
  value: ISearchResult[]
}

export default class SpfxSearchWithSliderProperty extends React.Component<ISpfxSearchWithSliderPropertyProps, ISearchResultsState> {

  constructor(props: ISpfxSearchWithSliderPropertyProps, state: ISearchResultsState)
  {
    super(props);

    /* Initialize the state object with our interface definition - empty values and loading = true */
    this.state = {
      searchResults: [] as ISearchResult[],
      loading: true
    };
  }

  /* React life-cycle: This method will be called on component load */
  public componentDidMount(): void {
    this.getSearchResults();
  }

  /* React life-cycle: This method will be called on component changed/update - eg. via PropertyPane panel */
  public componentDidUpdate(prevProps: ISpfxSearchWithSliderPropertyProps, prevState: ISearchResultsState, prevContext: any): void {
    if(this.props.count != prevProps.count || this.props.query != prevProps.query) {
      this.getSearchResults();
    }
  }

  private getSearchResults(): void {
    if(this.props.context.environment.type == EnvironmentType.Local)
    {
        this.getMockListData().then((response) => {
          this.setState((previousState: ISearchResultsState, curProps: ISpfxSearchWithSliderPropertyProps): ISearchResultsState => {
            previousState.loading = false;
            previousState.searchResults = response.value;
            return previousState;
          });
        });
    } else
    {
      //TODO: Search using REST API to SharePoint and parse into ISearchResult.
    }
  }

  private getMockListData(): Promise<IResult> {
    return MockHttpClient.get(this.props.context.pageContext.web.absoluteUrl).then(() => {
        const listData: IResult = {
            value:
            [

            ]
            };

            for(var i = 1; i<=this.props.count; i++) {
              listData.value.push(
                { title: 'Search Result ' + i, url: 'http://www.bjerkesolutions.dk', key: i + '', previewImageUrl: require('document-preview.png'), extension: 'docx' }
              );
            }

        return listData;
    }) as Promise<IResult>;
  }

  public render(): JSX.Element {



    const loading: JSX.Element = this.state.loading ? <div style={{margin: '0 auto'}}><Spinner label={'Loading...'} /></div> : <div/>;
    const results: JSX.Element[] = this.state.searchResults.map((res: ISearchResult, i: number) => {
      const iconUrl: string = `https://spoprod-a.akamaihd.net/files/odsp-next-prod_ship-2016-08-15_20160815.002/odsp-media/images/filetypes/32/${res.extension}.png`;
      return(
        <DocumentCard onClickHref={res.url} key={res.key}>
          <DocumentCardPreview
            previewImages={[
              {
                previewImageSrc: res.previewImageUrl,
                iconSrc: iconUrl,
                width: 318,
                height: 196,
                accentColor: '#ce4b1f'
              }
            ]}
            />
          <DocumentCardTitle title={res.title}/>
          <DocumentCardActivity
            activity='Created Sep 2, 2016'
            people={
            [
                { name: 'Kat Hudson', profileImageSrc: require('avatar-kat.png') }
            ]
            }
        />
        </DocumentCard>
      );
    });

    return (
      <div>
        {loading}
        {results}
      </div>
    );
  }
}
