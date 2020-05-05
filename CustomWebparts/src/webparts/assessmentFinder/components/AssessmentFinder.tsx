import * as React from 'react';
import styles from './AssessmentFinder.module.scss';
import { IAssessmentFinderProps } from './IAssessmentFinderProps';
import { SPDataOperations } from '../../../common/SPDataOperations';
import { PrimaryButton } from 'office-ui-fabric-react';

export interface IAssessmentFinderState{
  moduleAssessment:any[];
  assessmentStatus:string;
}
export default class AssessmentFinder extends React.Component<IAssessmentFinderProps, IAssessmentFinderState> {

  constructor(props) {
    super(props);

    this.state = {
      moduleAssessment:[],
      assessmentStatus:''
    };
  }

  public componentDidMount() {
    SPDataOperations.LOADCurrentUserAssessment(this.props.selectedList, this.props.assessmentList,1,this.props.context.pageContext.user.email,this.props.userAssessmentList).then((assessment) => {
      console.log(assessment);
      this.setState({moduleAssessment:assessment.assessmentData});
      if(assessment.assessmentData.length===0){
        this.setState({assessmentStatus:'Yay! you have no assessment pending.'});
      }
    });
  }
  public render(): React.ReactElement<IAssessmentFinderProps> {
    console.log(this.state);
    return (
      <div className={ styles.assessmentFinder }>
        <div className={ styles.container }>
          <div className={ styles.row }>
              {this.state.moduleAssessment.length > 0 &&
                <PrimaryButton href={this.props.description}>Start Assessment</PrimaryButton>
              }
              {this.state.moduleAssessment.length === 0 &&
                <p>{this.state.assessmentStatus}</p>
              }
          </div>
        </div>
      </div>
    );
  }
}
