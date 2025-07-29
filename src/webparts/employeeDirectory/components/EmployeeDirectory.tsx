import * as React from 'react';
// import styles from './EmployeeDirectory.module.scss';
import type { IEmployeeDirectoryProps } from './IEmployeeDirectoryProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { getSP } from '../pnp';
import { SPFI } from '@pnp/sp';
import { DetailsList,IColumn, Image, ImageFit } from '@fluentui/react';
interface IEmployeeItem{
  Title:String,
  Id:number,
  Email:String,
  Department:String,
  Designation:String,
  PhotoURL:Photo
} 
interface Photo{
  Description:String,
  Url:String
}
interface IEmployeeState{
  employees:IEmployeeItem[]
}
export default class EmployeeDirectory extends React.Component<IEmployeeDirectoryProps,IEmployeeState> {
  private sp : SPFI;
  private empTableColumns : IColumn[] = [
    {key:"col1",name:"Id",fieldName:"Id",minWidth:50,maxWidth:100,isResizable:true},
    {key:"col5",name:"Photo",fieldName:"PhotoURL",minWidth:50,maxWidth:100,isResizable:true,onRender:(item,index=0)=>{
      console.log("Item",item);
      return (
        <Image
        src={"https://picsum.photos/"+`${200+index}`}
        width={40}
        height={40}
        imageFit={ImageFit.cover}
        alt={item?.Title}
        styles={{root:{borderRadius:'50%'}}}
        />
      )
    }},
    {key:"col2",name:"Title",fieldName:"Title",minWidth:50,maxWidth:100,isResizable:true,isFiltered:true},
    {key:"col3",name:"Email",fieldName:"Email",minWidth:50,maxWidth:100,isResizable:true},
    {key:"col4",name:"Department",fieldName:"Department",minWidth:50,maxWidth:100,isResizable:true}
  ]
  constructor(props:IEmployeeDirectoryProps){
    super(props);
    this.state={
      employees:[]
    }
    this.sp=getSP(props.context);
  }
  public async componentDidMount(): Promise<void> {
    const employees =await this.getEmployeeDetails();
    this.setState({employees:employees});
    
  }
  private getEmployeeDetails =  async (): Promise<IEmployeeItem[]> => {
    //const employees:IEmployeeItem[] = 
    const items = await this.sp.web.lists.getByTitle('Employees').items
                        .select("Id","Title","Email","Department","Designation","PhotoURL")
                        .top(100)();
    //console.log("items",items)
    
    return items
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
    const {employees}= this.state;
    console.log("this",context,employees)
    return (
     <div>
      <h2>Employee Details</h2>
      <DetailsList
        items={employees}
        columns={this.empTableColumns}
      />
     </div>
    );
  }
}
