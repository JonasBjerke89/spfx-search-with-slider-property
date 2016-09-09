import { ISearchResult } from './../components/SpfxSearchWithSliderProperty';

export default class MockHttpClient {
  public static _results: ISearchResult[] = [];

  public static get(restURL: string, options?: any) : Promise<ISearchResult[]> {
    return new Promise<ISearchResult[]>((resolve) => {
      resolve(MockHttpClient._results);
    });
  }
}