import { ISPFxAdaptiveCard, BaseAdaptiveCardView, IActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'SubmitDataAceAdaptiveCardExtensionStrings';
// import {sp} from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";
import { sp } from "@pnp/sp/presets/all";

import { ISubmitDataAceAdaptiveCardExtensionProps,
  ISubmitDataAceAdaptiveCardExtensionState,
  QUICK_VIEW_CONFIRM_MENU_REGISTRY_ID
} from '../SubmitDataAceAdaptiveCardExtension';
import { PropertyPaneSlider } from '@microsoft/sp-property-pane';

export interface IChooseMenuQuickViewData {
  
  subTitle: string;
  title: string;
  description: string;
}

export class ChooseMenuQuickView extends BaseAdaptiveCardView<
  ISubmitDataAceAdaptiveCardExtensionProps,
  ISubmitDataAceAdaptiveCardExtensionState,
  IChooseMenuQuickViewData
> {
  public get data(): IChooseMenuQuickViewData {
    return {
      subTitle: strings.SubTitle,
      title: strings.Title,
      description: this.properties.description
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/ChooseMenuQuickViewTemplate.json');
  }

  public async onAction(action: IActionArguments | any) {
   
    // const list = sp.web.lists.getByTitle("LunchBooking");
   
     
    if (action.id == "Submit") {
      try{
      await sp.web.lists.getByTitle("LunchBooking").items.add({
      // this.setState({
        MainCourse: action.data.mainCourse,
        Dessert: action.data.dessert,
        Beverages: action.data.beverages, 
      // });
       
        

    })
    
      alert("Created Successfully");
  }
  catch (error) { console.log(error) }
      this.quickViewNavigator.push(QUICK_VIEW_CONFIRM_MENU_REGISTRY_ID);
    }
  }
}