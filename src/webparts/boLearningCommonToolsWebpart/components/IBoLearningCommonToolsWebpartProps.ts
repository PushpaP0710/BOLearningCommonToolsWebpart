import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IBoLearningCommonToolsWebpartProps {
  ListTitle: string;  
  SiteUrl: string; 
  context:  WebPartContext; 
}

export interface IBoLearningCommonToolsWebpartState {
  items:[    
    {    
      "Title": "",   
      "NavigateURL": "",   
      "Logo":{
          "Url":"",
          "Description":"",
          }  
    }];
}
