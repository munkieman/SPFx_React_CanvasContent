import * as React from 'react';
import styles from './CanvasContent.module.scss';
import type { ICanvasContentProps } from './ICanvasContentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';

//let webPartTitle:string = "";
//let grouptitle1:string="";
//let grouptitle2:string="";
//let grouptitle3:string="";
//let grouptitle4:string="";
//let grouptitle5:string="";

let panelHTML: any[]=[];

export interface IStates {
}

export default class CanvasContent extends React.Component<ICanvasContentProps, IStates, {}> {

  //constructor(props){
  //  super(props);

  //  this.state = {};
  //}

  public componentDidMount(): void {
    console.log("component did mount");
    //this._getWebPartData();    
  }

  public componentWillMount(): void {
    console.log("component will mount");
    //this._getWebPartData();    
  }

  public componentDidUpdate(): void {
    console.log("component did update");
    //this._getWebPartData();    
  }

  public componentWillUnmount(): void {
    console.log("component will unmount");
    //this._getWebPartData();    
  }

  public render(): React.ReactElement<ICanvasContentProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    this._renderWebPartDataAsync();    
    console.log("render data",panelHTML);

    return (
      <section className={`${styles.canvasContent} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        
        <div id="spContainer"></div>

        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
      </section>
    );    
  }

  private _renderWebPartDataAsync(): void {
    this._getWebPartData()
      .then(async (response:any) => {
        this._renderWebPartData(response);
      });
  }

  private async _getWebPartData() {
    const endpoint = `${this.props.siteUrl}/_api/sitepages/pages(1)?$select=CanvasContent1&expand=CanvasContent1`;
    const rawResponse = await this.props.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
    const jsonResponse = await rawResponse.json();
    const jsonCanvasContent = jsonResponse.CanvasContent1;
    const parseCanvasContent = JSON.parse(jsonCanvasContent);
    return parseCanvasContent;
/*
    return await this.props.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
      .then(async (response) => {
        if(response.ok){
          console.log("response json",response.json());
          return await response.json();
        }
      })
*/
  }

  public _renderWebPartData (items:any): void {
    
    console.log("items",items);
    
    let html: string="";
    for(const item of items){
      console.log("item",item.id);
      //html += `<div><h1>webpart title: ${item.webPartData.title}</h1></div>`;
    }
    //html += '<div><h1>webpart title: </h1></div>';

    //const listContainer: Element = document.querySelector('#spListContainer')!;
    document.querySelector('#spContainer')!.innerHTML = html;
  }

  /*
  private async getWebPartData(): Promise<void>{
    console.log("Fetching CanvasContent1 WebPartData Title.");
    const endpoint = `${this.props.siteUrl}/_api/sitepages/pages(1)?$select=CanvasContent1&expand=CanvasContent1`;
    const rawResponse = await this.props.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
    const jsonResponse = await rawResponse.json();
    const jsonCanvasContent = jsonResponse.CanvasContent1;
    const parseCanvasContent = JSON.parse(jsonCanvasContent);

    console.log("canvascontent parse",parseCanvasContent);
    for(const item of parseCanvasContent){
      console.log("webPartData Title",item.webPartData.title);
      let itemTitle : string = item.webPartData.title;
      //let itemGroup1 : string = item.webPartData.properties.Group1Title;
      //let itemGroup2 : string = item.webPartData.properties.Group2Title;
      //let itemGroup3 : string = item.webPartData.properties.Group3Title;
      //let itemGroup4 : string = item.webPartData.properties.Group4Title;
      //let itemGroup5 : string = item.webPartData.properties.Group5Title;

      if(item.webPartData.title === "Important Links"){
        //webPartTitle=itemTitle;
        //grouptitle1=itemGroup1;
        //grouptitle2=itemGroup2;
        //grouptitle3=itemGroup3;
        //grouptitle4=itemGroup4;
        //grouptitle5=itemGroup5;
        panelHTML+=`<div><h1>webpart title: ${itemTitle}</h1></div>`;

//        <h1>webpart title: {webPartTitle}</h1>
//        <h1>grouptitle1: {grouptitle1}</h1>
//        <h1>grouptitle2: {escape(grouptitle2)}</h1>
//        <h1>grouptitle3: {escape(grouptitle3)}</h1>
//        <h1>grouptitle4: {escape(grouptitle4)}</h1>
//        <h1>grouptitle5: {escape(grouptitle5)}</h1>

        break;          
      }
    }
  } 
    */   
}