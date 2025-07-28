import * as React from 'react';
// import styles from './EmployeeDirectory.module.scss';
import type { IEmployeeDirectoryProps } from './IEmployeeDirectoryProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { getSP } from '../pnp';
import { SPFI } from '@pnp/sp';
interface IEmployeeItem{
  Title:String,
  Id:number
}
interface IEmployeeState{
  employees:IEmployeeItem[]
}
export default class EmployeeDirectory extends React.Component<IEmployeeDirectoryProps,IEmployeeState> {
  private sp : SPFI;
  constructor(props:IEmployeeDirectoryProps){
    super(props);
    this.state={
      employees:[]
    }
    this.sp=getSP(props.context);
  }
  public componentDidMount(): void {
    this.getEmployeeDetails();
  }
  private getEmployeeDetails =  async (): Promise<void> => {
    //const employees:IEmployeeItem[] = 
    const items = await this.sp.web.lists.getByTitle('Employees').items.select("Id","Title","Email","Department","Designation","PhotoURL");
    console.log(items);
    //return employees
  }
  public render(): React.ReactElement<IEmployeeDirectoryProps> {
    
    const {
      // description,
      // isDarkTheme,
      // environmentMessage,
      // hasTeamsContext,
      // userDisplayName,
      context
    } = this.props;
    console.log("this",context)
    return (
     <div>
      <h2>Employee Details</h2>
      <ul>
        {}
      </ul>
     </div>
    );
  }
}
