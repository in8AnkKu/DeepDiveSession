import { sp, PermissionKind } from '@pnp/sp';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
//Polyfill to fix IE issues
import '@pnp/polyfill-ie11';
import Assessment from '../webparts/assessment/components/Assessment';
import * as pnp from "sp-pnp-js";

export class SPDataOperations {

  /**
   * Gets the available Choices in the Module Choice field
   * @param lists The list for which the fields of type Managed Metadata need to be retrieved
   */
  public static async LOADCurrentUserTraining(lists: string, userId:number): Promise<any> {
    let selectedTraining: any[] = [];
    let selectedTrainingObject:any[] = [];
    try {
      let userData = await sp.web.lists.getByTitle(lists).items.select('myTraining/Id,myTraining/ModuleCalc').expand('myTraining').filter(`Id eq `+userId).get();
      userData[0].myTraining.map((training) =>{
        selectedTraining.push(training.Id);
        selectedTrainingObject.push({'Module':training.ModuleCalc,'Id':training.Id});
      });
    } catch (error) {
      console.log(error.message);
    }

    let allselectedTraining:any = {'selectedTraining':selectedTraining};
    let allselectedTrainingObject:any = {'selectedTrainingObject':selectedTrainingObject};
    selectedTraining = {...allselectedTraining,...allselectedTrainingObject};
    return selectedTraining;
  }

    /**
   * Gets the available sub module
   *
   * @param lists The list for which the fields of type Managed Metadata need to be retrieved
   * @param module
   */
  public static async LOADSubModuleData(lists: string): Promise<any> {
    let allData: any;
    let selectedTraining: any;
    let moduleData:any[] = [];
    let subModuleData:any = {};
    let trainingData:any = {};
    let trainingIDs:any = {};
    try {
      selectedTraining = await this.LOADCurrentUserTraining('testList',1);
      allData = await sp.web.lists.getById(lists).items.select('Id,Title,Module,SubModule,TrainingPath').get();
      allData.map((field) =>{
        if(moduleData.indexOf(field.Module) === -1){
          moduleData.push(field.Module);
          subModuleData[field.Module] = [];
          trainingIDs[field.Module] = [];
        }
      });
      allData.map((field) =>{
        if(subModuleData[field.Module].indexOf(field.SubModule) === -1){
          subModuleData[field.Module].push(field.SubModule);
          trainingData[field.SubModule] = [];
        }
      });
      allData.map((field) =>{
        if(trainingData[field.SubModule].indexOf(field) === -1){
          trainingData[field.SubModule].push(field);
          trainingIDs[field.Module].push(field.Id);
        }
      });
    } catch (error) {
      console.log(error.message);
    }

    let allSelectedTraining = {'selectedTraining':selectedTraining.selectedTraining};
    let allModuleData:any = {'module':moduleData};
    let allSubModuleData:any = {'subModule':subModuleData};
    let allTrainingData:any = {'trainingData':trainingData};
    let allTrainingIds:any = {'trainingIds':trainingIDs};
    allData = {...allModuleData,...allSubModuleData,...allTrainingData,...allSelectedTraining,...allTrainingIds};
    return allData;
  }
  /**
   * Check if the current user has requested permissions on a list
   *
   * @param lists The list on which user permission needs to be checked
   * @param permission The permission kind for which user needs to be authorized
   */
  public static async CHECKUSERPERMISSIONS(lists: string, permission: PermissionKind): Promise<boolean> {
    let perms: boolean = false;
    try {
      let effectiveBasePermissions = await sp.web.lists.getById(lists).effectiveBasePermissions.get();
      perms = await sp.web.lists.getById(lists).hasPermissions(effectiveBasePermissions, permission);
    } catch (error) {
      console.log(error.message);
    }
    return perms;
  }

  /**
   * Check if the current user has requested permissions on a list
   *
   * @param lists The list on which user permission needs to be checked
   * @param ids The permission kind for which user needs to be authorized
   * @param itemId
   * @param pageContext
   * @param props
   * @param CompletedModule
   * @param userAssessmentList
   */
  public static async UpdateTrainings(lists: string, trainingIds:any[],itemId:number,pageContext:any,props:any,CompletedModule?:string,userAssessmentList?: string){
    let SPDATA = await this.getListItemEntityType(lists);
    const body: string = JSON.stringify({
      '__metadata': { 'type': SPDATA },
      'myTrainingId': {
        'results': trainingIds 
     }
    });
  
    props.spHttpClient.post(`${pageContext.web.absoluteUrl}/_api/web/lists/getbyid('${lists}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      {
      headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=verbose',
      'odata-version': '',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE'
      },
      body: body
      })
      .then((response: SPHttpClientResponse): void => {
        if(CompletedModule !== "" && CompletedModule !== undefined){
          this.AssignModuleAssessment(userAssessmentList,CompletedModule,props);
        } else {
          window.location.reload();
        }
      }, (error: any): void => {
        console.log(error);
      });
  }
  
    /**
   * Gets the available Choices in the Module Choice field
   * @param lists The list for which the fields of type Managed Metadata need to be retrieved
   * @param assessmentlist
   * @param totalQuestion
   * @param userEmail
   * @param userAssessmentList
   */
  public static async LOADCurrentUserAssessment(lists: string, assessmentlist:string, totalQuestion:any, userEmail:string, userAssessmentList: string): Promise<any> {
    let selecedModule: any;
    let assessments: any[] = [];
    let correctAnswer: any = {};
    let userAnswer: any = {};
    let assessmentAttempt: any = {};
    try {
     
      let userData:any = await sp.web.lists.getById(lists).items.select('CompletedModule,EmployeeID/EMail').expand('EmployeeID').filter(`EmployeeID/EMail eq '`+userEmail+`'`).get();
      selecedModule = userData.length > 0 ? userData[0].CompletedModule : "";

      if(selecedModule !== ""){
        assessmentAttempt = await this.GetAssessmentStatus(userAssessmentList,userEmail);

        if(assessmentAttempt.totalAttempt === 0 || (assessmentAttempt.assessmentStatus === 'Fail' && assessmentAttempt.totalAttempt < 3)){
          assessments = await sp.web.lists.getById(assessmentlist).items.select('Id,Title,OData__x004f_pt1,OData__x004f_pt2,OData__x004f_pt3,OData__x004f_pt4,Answer').filter(`Module eq '`+encodeURIComponent(selecedModule)+`'` ).get();
          assessments.sort((a, b) => {return 0.5 - Math.random();});
          
          const selectedItems = assessments.slice(0, +totalQuestion).map(item => {
            correctAnswer[item.Id] = item.Answer;
            userAnswer[item.Id] = "";
            return item;
          });
          assessments = selectedItems;
        }

      }
    } catch (error) {
      console.log(error.message);
    }
    let assessmentModule:any = {'assessmentModule':selecedModule};
    let assessmentData:any = {'assessmentData':assessments};
    let assessmentAnswer:any = {'correctAnswer':correctAnswer};
    let AssessmentQuestion:any = {'userAnswer':userAnswer};
    let assessmentTotalAttempt:any = {'totalAttempt':assessmentAttempt};
    return {...assessmentModule,...assessmentData,...assessmentAnswer,...AssessmentQuestion,...assessmentTotalAttempt};
  }

  /**
   * Check if the current user has requested permissions on a list
   *
   * @param lists The list on which user permission needs to be checked
   * @param userEmail The permission kind for which user needs to be authorized
   */

  public static async GetAssessmentStatus(lists: string, userEmail: string){
    let assessmentAttemptData = await sp.web.lists.getById(lists).items.select('*,EmployeeID/EMail').expand('EmployeeID').orderBy('Created', false).top(1).filter(`EmployeeID/EMail eq '`+userEmail+`'`).get();
    return assessmentAttemptData.length > 0 ? {attemptId:assessmentAttemptData[0].Id,totalAttempt:assessmentAttemptData[0].Attempted,assessmentStatus:assessmentAttemptData[0].AssessmentStatus || '',assessmentAllData:assessmentAttemptData[0]} : {attemptId:0,totalAttempt:0,assessmentStatus:'',assessmentAllData:{}};
  }
/**
   * Check if the current user has requested permissions on a list
   *
   * @param lists The list on which user permission needs to be checked
   * @param ids The permission kind for which user needs to be authorized
   * @param itemId
   * @param pageContext
   * @param props
   */
  public static async UpdateAssessmentStatus(lists: string, module: string, status: string, totalAttemptData:any,props:any, correctQuestion:number,score:number,totalQuestion:number){
    let totalAttempt:number = totalAttemptData.totalAttempt+1;
    let SPDATA = await this.getListItemEntityType(lists);
    const body: string = JSON.stringify({
      '__metadata': { 'type': SPDATA },
      'Attempted': totalAttempt,
      'AssessmentStatus':status,
      'totalQuestion': totalQuestion,
      'passingScore': +props.passingScore,
      'correctQuestion': correctQuestion,
      'score':score.toFixed(2)
    });
  
    props.context.spHttpClient.post(`${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbyid('${lists}')/items(${totalAttemptData.attemptId})`,
      SPHttpClient.configurations.v1,
      {
      headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=verbose',
      'odata-version': '',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE'
      },
      body: body
      })
      .then(async (response: SPHttpClientResponse): Promise<void> => {
        if(totalAttempt===3 && status==="Fail"){
          let selectedTraining = await this.LOADCurrentUserTraining(props.userTrainingList,1);
          let selectedTrainingObject = selectedTraining.selectedTrainingObject;
          let updatedTrainingId:any = [];
          selectedTrainingObject.map((val) => {
            if(val.Module != module){
              updatedTrainingId.push(val.Id);
            }
          });
          const updateTraining = await this.UpdateTrainings(props.userTrainingList,updatedTrainingId,1,props.context.pageContext,props.context);
        } else {
          window.location.href = window.location.pathname+"?assessment=true";
        }
      }, (error: any): void => {
        console.log(error);
      });
  }

  public static async AssignModuleAssessment(lists: string, module: string, props:any){
    let userDetails = await this.spLoggedInUserDetails(props);
    let SPDATA = await this.getListItemEntityType(lists);
    const body: string = JSON.stringify({
      '__metadata': { 'type': SPDATA },
      'Title': module,
      'EmployeeIDId':userDetails.Id
    });
  
    props.spHttpClient.post(`${props.pageContext.web.absoluteUrl}/_api/web/lists/getbyid('${lists}')/items`,
      SPHttpClient.configurations.v1,
      {
      headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=verbose',
      'odata-version': ''
      },
      body: body
      })
      .then(async (response: SPHttpClientResponse): Promise<void> => {
        window.location.href = window.location.pathname+"?assessment=true";
      });
  }

/*Get Current Logged In User*/  
public static async spLoggedInUserDetails(ctx: any): Promise<any>{  
  try {  
      const web = new pnp.Web(ctx.pageContext.site.absoluteUrl);  
      return await web.currentUser.get();          
    } catch (error) {  
      console.log("Error in spLoggedInUserDetails : " + error);  
    }      
  } 
/**
   * Check if the current user has requested permissions on a list
   * @param listId The list on which user permission needs to be checked
   */
  public static async getListItemEntityType(listId: string){
    let entityType:any;
    try {
      entityType = await sp.web.lists.getById(listId).getListItemEntityTypeFullName();
    } catch(error){
      console.log('SPDataOperations.getListItemEntityType' + error);
    }
    return entityType;
  }
}