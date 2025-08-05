import * as React from 'react';
// import styles from './EmployeeDirectory.module.scss';
import type { IEmployeeDirectoryProps } from './IEmployeeDirectoryProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { getSP } from '../pnp';
import { SPFI } from '@pnp/sp';
import { DetailsList, IColumn, IconButton, Image, ImageFit, Panel, PrimaryButton, SearchBox } from '@fluentui/react';
interface IEmployeeItem {
  Title: String,
  Id: number,
  Email: String,
  Department: String,
  Designation: String,
  PhotoURL: Photo
}
interface Photo {
  Description: String,
  Url: String
}
interface IEmployeeState {
  employees: IEmployeeItem[],
  isPanelOpen?: boolean
}
export default class EmployeeDirectory extends React.Component<IEmployeeDirectoryProps, IEmployeeState> {
  private sp: SPFI;
  private empTableColumns: IColumn[] = [
    { key: "col1", name: "Id", fieldName: "Id", minWidth: 50, maxWidth: 100, isResizable: true },
    {
      key: "col5", name: "Photo", fieldName: "PhotoURL", minWidth: 50, maxWidth: 100, isResizable: true, onRender: (item, index = 0) => {
        console.log("Item", item);
        return (
          <Image
            src={"https://picsum.photos/" + `${200 + index}`}
            width={40}
            height={40}
            imageFit={ImageFit.cover}
            alt={item?.Title}
            styles={{ root: { borderRadius: '50%' } }}
          />
        )
      }
    },
    { key: "col2", name: "Title", fieldName: "Title", minWidth: 50, maxWidth: 100, isResizable: true, isFiltered: true },
    { key: "col3", name: "Email", fieldName: "Email", minWidth: 50, maxWidth: 100, isResizable: true },
    { key: "col4", name: "Department", fieldName: "Department", minWidth: 50, maxWidth: 100, isResizable: true }
  ]
  constructor(props: IEmployeeDirectoryProps) {
    super(props);
    this.state = {
      employees: [],
      isPanelOpen: false
    }
    this.sp = getSP(props.context);
  }
  public async componentDidMount(): Promise<void> {
    const employees = await this.getEmployeeDetails();
    this.setState({ employees: employees });

  }
  private getEmployeeDetails = async (): Promise<IEmployeeItem[]> => {
    //const employees:IEmployeeItem[] = 
    const items = await this.sp.web.lists.getByTitle('Employees').items
      .select("Id", "Title", "Email", "Department", "Designation", "PhotoURL")
      .top(100)();
    //console.log("items",items)

    return items
  }
  private openPanel = (): void => {
    this.setState({ isPanelOpen: true });
  }
  private closePanel = (): void => {
    this.setState({ isPanelOpen: false });
  }
  public render(): React.ReactElement<IEmployeeDirectoryProps> {

    // const {
    //   // description,
    //   // isDarkTheme,
    //   // environmentMessage,
    //   // hasTeamsContext,
    //   // userDisplayName,
    //   context
    // } = this.props;
    const { employees, isPanelOpen } = this.state;
    // console.log("this",context,employees)
    return (
      <div>
        <Panel
          isOpen={isPanelOpen}
          onDismiss={() => { this.closePanel() }}
          headerText="Filter Employees"
          closeButtonAriaLabel="Close"
          isLightDismiss={true}
          isFooterAtBottom={true}
          // onRenderBody={() => (
          //   <div style={{padding:'20px'}}>
          //     <Dropdown
          //   </div>
          // )}
          onRenderFooterContent={() => (
            <div style={{ display: 'flex', justifyContent: 'space-between' }}>
              <PrimaryButton
                text="Apply"
                onClick={() => { this.closePanel() }}
                style={{ marginRight: '10px' }} />
              {/* <IconButton
                iconProps={{ iconName: 'Cancel' }}
                title="Cancel"
                ariaLabel="Cancel"
                onClick={() => { this.closePanel() }}
              /> */}
            </div>
          )}
        // styles={{main:{width:'400px'}}}
        >
        </Panel>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <h2>Employee Directory</h2>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <IconButton
              iconProps={{ iconName: 'Add' }}
              title="Add Employee"
              ariaLabel="Add Employee"
              onClick={() => { }}
            />
            <IconButton
              iconProps={{ iconName: 'Filter' }}
              title="Filter"
              ariaLabel="Filter"
              onClick={() => { this.openPanel() }}
            />
            <SearchBox
              placeholder="Search Employee"
              onSearch={(value: string) => { }}
              onChange={(ev, value) => { }}
              styles={{ root: { width: 200 } }}
              ariaLabel="Search Employee"
              iconProps={{ iconName: 'Search' }}
              clearButtonProps={{ ariaLabel: 'Clear search' }}
              onClear={(ev) => { }}
              underlined={true}
            />
          </div>
        </div>
        <div style={{ height: '500px', overflow: 'auto' }}>
          <DetailsList
            items={employees}
            columns={this.empTableColumns}
            setKey="set"
            layoutMode={0} // 0 is fixedColumns
            selectionPreservedOnEmptyClick={true}
            compact={true}
            isHeaderVisible={true}
          // onRenderRow={(props,defaultRender)=>{
          //   return defaultRender(props);
          // }}
          />
        </div>
      </div>
    );
  }
}
