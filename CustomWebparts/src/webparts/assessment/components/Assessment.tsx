import * as React from 'react';
import styles from './Assessment.module.scss';
import { IAssessmentProps } from './IAssessmentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import { DisplayMode } from '@microsoft/sp-core-library';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { SPDataOperations } from '../../../common/SPDataOperations';
import {DefaultButton, PrimaryButton} from 'office-ui-fabric-react/lib';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
export interface IAssessmentState{
  module:any;
  moduleAssessment:any[];
  userAnswer:any;
  correctAnswer:any;
  totalAttemptData:any;
  assessmentModule: any;
  isOpenPromt:boolean;
  assessmentParm:boolean;
}

export default class Assessment extends React.Component<IAssessmentProps, IAssessmentState> {

  constructor(props) {
    super(props);

    this.state = {
      module:'',
      moduleAssessment:[],
      userAnswer:{},
      correctAnswer:{},
      totalAttemptData:{},
      assessmentModule:{},
      isOpenPromt:true,
      assessmentParm:true
    };

    this.onConfigure = this.onConfigure.bind(this);
    this._onChange = this._onChange.bind(this);
    this.submittedAssessment = this.submittedAssessment.bind(this);
  }

  public componentDidMount() {
    this.renderAssessmentModule();
    let queryParms = new UrlQueryParameterCollection(window.location.href);
    let assessmentParm:any = queryParms.getValue("assessment");
    if(assessmentParm === true || assessmentParm === 'true'){
      this.setState({assessmentParm:false});
    }
  }

  public componentDidUpdate(prevProps: IAssessmentProps) {
    if(prevProps.userTrainingList !== this.props.userTrainingList || prevProps.assessmentList !== this.props.assessmentList || prevProps.totalQuestion !== this.props.totalQuestion){
      this.renderAssessmentModule();
    }
  }

  public renderAssessmentModule(){
    SPDataOperations.LOADCurrentUserAssessment(this.props.userTrainingList, this.props.assessmentList,this.props.totalQuestion,this.props.context.pageContext.user.email, this.props.userAssessmentList).then((allTrainigs) => {
      this.setState({module:allTrainigs.assessmentModule,moduleAssessment:allTrainigs.assessmentData,correctAnswer:allTrainigs.correctAnswer,userAnswer:allTrainigs.userAnswer,totalAttemptData:allTrainigs.totalAttempt,assessmentModule:allTrainigs.totalAttempt.assessmentAllData});
    });
  }

  private onConfigure(): void {
    this.props.context.propertyPane.open();
  }

  public _onChange(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): void {
    let userGivenAnswer = this.state.userAnswer;
    let selectedAnswer = option.key.split("_");
    userGivenAnswer[selectedAnswer[0]] = selectedAnswer[1];
    this.setState({userAnswer:userGivenAnswer});
  }

  public submittedAssessment(){
    const passingScore:any = +this.props.passingScore;
    const userAnswer:any = this.state.userAnswer;
    const correctAnswer:any = this.state.correctAnswer;
    const totalAttemptData:any = this.state.totalAttemptData;
    let correctAnsNo:number = 0;
    let totalQuestion:number = 0;
    
    Object.keys(correctAnswer).map((ans) =>{
      if(correctAnswer[ans]===userAnswer[ans]){
        correctAnsNo++;
      }
      totalQuestion++;
    });
  
    const totalPercentage:number = (correctAnsNo*100)/totalQuestion;
    const userStatus:any = (totalPercentage-passingScore) > 0 ? "Pass" : "Fail";
    SPDataOperations.UpdateAssessmentStatus(this.props.userAssessmentList,this.state.module,userStatus,totalAttemptData,this.props,correctAnsNo,totalPercentage,totalQuestion).then((allTrainigs) => {
    });
  }

  public render(): React.ReactElement<IAssessmentProps> {
    console.log(this.state);
    if (this.props.configured) {
      let assessmentAllData = this.state.assessmentModule;
    return (
      <div className={ styles.assessment }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            {this.state.moduleAssessment.map((item,i) => {
              const options: IChoiceGroupOption[] = [
                { key: item.Id+'_Opt1', text: item.OData__x004f_pt1 },
                { key: item.Id+'_Opt2', text: item.OData__x004f_pt2 },
                { key: item.Id+'_Opt3', text: item.OData__x004f_pt3 },
                { key: item.Id+'_Opt4', text: item.OData__x004f_pt4 }
              ];              
              return (<div className={styles.questionRow}>
                <ChoiceGroup key={item.Id} options={options} label={"Q."+(i+1)+" "+item.Title} onChange={this._onChange} />
              </div>);
            })
            }
            {this.state.moduleAssessment.length !== 0 &&
            <div>
 <Dialog
    hidden={this.state.isOpenPromt}
    dialogContentProps={{
      type: DialogType.normal,
      title: 'Confirm!',
      closeButtonAriaLabel: 'Close',
      subText: 'Do you want to submit the current module?'
    }}
  >
    <DialogFooter>
      <PrimaryButton onClick={this.submittedAssessment} text="OK" />
      <DefaultButton onClick={() => { this.setState({isOpenPromt: true});}}  text="Cancel" />
    </DialogFooter>
  </Dialog>
              <PrimaryButton onClick={() => { this.setState({isOpenPromt: false});}}>Submit</PrimaryButton>
            </div>
            }
            {(this.state.moduleAssessment.length === 0 && this.state.assessmentModule.Attempted) &&
            <div style={{textAlign:"center"}}>
              <img style={{width:'auto'}} src="/sites/" />
              <h2>You have no assessment pending.</h2>
            </div>
            }
{assessmentAllData.Attempted > 0 &&
  <Dialog
    hidden={this.state.assessmentParm}
    dialogContentProps={{
      type: DialogType.largeHeader,
      title:'Your Assessment Score'
    }}
    modalProps={{
      isBlocking: true
    }}
    containerClassName={styles.dialogContainer}
  >
    <div className={styles.container}>
    <div className={styles.row}>
      <table>
        <tr>
          <th>Module</th>
          <td>{assessmentAllData.Title}</td>
        </tr>
        <tr>
          <th>Total Question</th>
          <td>{assessmentAllData.totalQuestion}</td>
        </tr>
        <tr>
          <th>Correct Question</th>
          <td>{assessmentAllData.correctQuestion}</td>
        </tr>
        <tr>
          <th>Passing Score (%)</th>
          <td>{assessmentAllData.passingScore}%</td>
        </tr>
        <tr>
          <th>Your Score (%)</th>
          <td>{assessmentAllData.score}%</td>
        </tr>
        <tr>
          <th>Total Attempt</th>
          <td>{assessmentAllData.Attempted} of 3</td>
        </tr>
        <tr>
          <th>Status</th>
          <td><b style={{color:assessmentAllData.AssessmentStatus==='Pass'?'Green':'Red'}}>{assessmentAllData.AssessmentStatus}</b></td>
        </tr>
      </table>
      <MessageBar
        messageBarType={assessmentAllData.AssessmentStatus==='Pass'?MessageBarType.success:MessageBarType.severeWarning}
        isMultiline={true}
        >
          {assessmentAllData.AssessmentStatus==='Pass' &&
            <div><b>Congratulations</b>, you have passed this assessment!<br/>
              <i>Note: Please complete your pending training if any.</i></div>
          }
          {(assessmentAllData.AssessmentStatus==='Fail' && assessmentAllData.Attempted < 3) &&
            <div>After 3 failed attempt you will have to complete the training again for this module.</div>
          }
          {(assessmentAllData.AssessmentStatus==='Fail' && assessmentAllData.Attempted===3) &&
          <div>You have 3 failed attempts at this assessment. Please retake the training for this module prior to attempting the assessment again.</div>
          }
        </MessageBar>

      </div>
    </div>
    <DialogFooter>
      <PrimaryButton onClick={() => { this.setState({assessmentParm: true});}}  text="OK" />
    </DialogFooter>
  </Dialog>
}
          </div>

        </div>
      </div>
    );
    } else {
      return (
        <Placeholder iconName='Edit'
          iconText='Configure your web part'
          description='Please configure the web part.'
          buttonLabel='Configure'
          hideButton={this.props.displayMode === DisplayMode.Read}
          onConfigure={this.onConfigure} />
      );
    }
  }
}
