import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { AadHttpClientFactory, AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";

export default class FunctionService {
  private aadHttpClientFactory: AadHttpClientFactory;
  private client: AadHttpClient;

  public static readonly serviceKey: ServiceKey<FunctionService> =
    ServiceKey.create<FunctionService>('react-site-secure-function-call-smpl', FunctionService);

  constructor(serviceScope: ServiceScope) {  
    serviceScope.whenFinished(async () => {
      this.aadHttpClientFactory = serviceScope.consume(AadHttpClientFactory.serviceKey);      
    });
  }

  public async setNewSiteDescretion(siteUrl: string, siteDescreption: string): Promise<any[]> {
    this.client = await this.aadHttpClientFactory.getClient('api://mmospfxsecsamplefunction.azurewebsites.net/0a8dfbc9-0423-495b-a1e6-1055f0ca69c2');
    /* const requestBody = {      
      URL: siteUrl,
      Descreption: siteDescreption
    }; */
    const requestUrl = `http://localhost:7241/api/SiteFunction?URL=${siteUrl}&Descreption=${siteDescreption}`;
    return this.client
      .get(requestUrl, AadHttpClient.configurations.v1
              /* { 
                body: JSON.stringify(requestBody) 
              } */
            )   
      .then((response: HttpClientResponse) => {
        return response.json();
      });
    
  }

}