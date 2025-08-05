import * as React from 'react';
// import styles from './EmployeeDirectory.module.scss';
import type { IEmployeeDirectoryProps } from './IEmployeeDirectoryProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { getSP } from '../pnp';
import { SPFI } from '@pnp/sp';
import { DetailsList, IColumn, IconButton, Image, ImageFit, Panel, PrimaryButton, SearchBox } from '@fluentui/react';
import { debounce } from '../utils';
interface IEmployeeItem {
  Title: string,
  Id: number,
  Email: string,
  Department: string,
  Designation: string,
  PhotoURL: Photo
}
interface Photo {
  Description: string,
  Url: string
}
interface IEmployeeState {
  employees: IEmployeeItem[],
  isPanelOpen?: boolean,
  // searchText?: string,
  filterOptions?: filterItem[]
}
interface filterItem{
  key:string,
  value?:string
  // type:string
}
export default class EmployeeDirectory extends React.Component<IEmployeeDirectoryProps, IEmployeeState> {
  private sp: SPFI;
  private empTableColumns: IColumn[] = [
    { key: "col1", name: "Id", fieldName: "Id", minWidth: 50, maxWidth: 100, isResizable: true },
    {
      key: "col5", name: "Photo", fieldName: "PhotoURL", minWidth: 50, maxWidth: 100, isResizable: true, onRender: (item, index = 0) => {
        // console.log("Item", item);
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
  private batchSize: number = 10; // Number of items to fetch in each batch
  private skipCount: number = 0; // Counter to keep track of skipped items
  private employeeListRef: React.RefObject<HTMLDivElement> = React.createRef<HTMLDivElement>();
  constructor(props: IEmployeeDirectoryProps) {
    super(props);
    this.state = {
      employees: [],
      isPanelOpen: false,
      // searchText: undefined
    }
    this.sp = getSP(props.context);
  }
  public async componentDidMount(): Promise<void> {
    const employees = await this.getEmployeeDetails();
    this.setState({ employees: employees });
    // Adding scroll event listener to the employee list
    if (this.employeeListRef.current) {
      this.employeeListRef.current.addEventListener('scroll', this.handleNativeScroll);
    }
  }
  async componentDidUpdate(prevProps: Readonly<IEmployeeDirectoryProps>, prevState: Readonly<IEmployeeState>, snapshot?: any): Promise<void> {
    if (prevState.filterOptions !== this.state.filterOptions) {
      console.log("Search Text Changed", this.state.filterOptions, typeof this.state.filterOptions);
      this.clearSkipCount(); // Reset skip count when filter options change
      const newEmployees = await this.getEmployeeDetails();
      this.setState((prevState) => ({
        employees: [...prevState.employees, ...newEmployees]
      }));
      // this.filterEmployees();
    }
  }
  componentWillUnmount(): void {
    // Removing scroll event listener to prevent memory leaks
    if (this.employeeListRef.current) {
      this.employeeListRef.current.removeEventListener('scroll', this.handleNativeScroll);
    }
  }
  private handleNativeScroll = async (event: Event): Promise<void> => {
    const target = event.target as HTMLDivElement;
    // Check if the user has scrolled to the bottom of the list
    console.log("Scroll Event",target.scrollHeight - target.scrollTop,target.clientHeight, target.scrollHeight, target.scrollTop);
    if (target.scrollHeight - target.scrollTop <= target.clientHeight) {
      // Load more employees
      this.skipCount += this.batchSize; // Increment skip count
      const newEmployees = await this.getEmployeeDetails();
      this.setState((prevState) => ({
        employees: [...prevState.employees, ...newEmployees]
      }));
    }
  }
  private getEmployeeDetails = async (): Promise<IEmployeeItem[]> => {
    //const employees:IEmployeeItem[] = 
    const {filterOptions} = this.state;
    let items: IEmployeeItem[] = [];
    if(filterOptions && filterOptions.length > 0){
      // console.log("Filter Options", filterOptions);
      const filterQuery = filterOptions.filter(value => value).map(option => option.value).join(' and ');
      items = await this
                        .sp
                        .web
                        .lists
                        .getByTitle('Employees')
                        .items
                        .select("Id", "Title", "Email", "Department", "Designation", "PhotoURL")
                        .filter(filterQuery)
                        .orderBy("Id", true)
                        .skip(this.skipCount)
                        .top(this.batchSize)(); // Fetching only the specified batch size
          
    }else{
      items = await this
                        .sp
                        .web
                        .lists
                        .getByTitle('Employees')
                        .items
                        .select("Id", "Title", "Email", "Department", "Designation", "PhotoURL")
                        .orderBy("Id", true)
                        .skip(this.skipCount)
                        .top(this.batchSize)() // Fetching only the specified batch size
    }

    return items
  }
  private openPanel = (): void => {
    this.setState({ isPanelOpen: true });
  }
  private closePanel = (): void => {
    this.setState({ isPanelOpen: false });
  }
  private clearSkipCount = (): void => {
    this.skipCount = 0; // Reset skip count
   // this.setState({ employees: [] }); // Clear the current employee list
    this.employeeListRef.current?.scrollTo({ top: 0 }); // Scroll to the top of the list
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
    const { employees, isPanelOpen} = this.state;
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
              onSearch={debounce((newValue:string) => {
                  console.log("onSearch Text", newValue); 
                  this.setState((prevState) => ({
                    // searchText: newValue == null || newValue == "" ? undefined : newValue,
                    employees: [],
                    filterOptions: [
                      ...(prevState.filterOptions || []),
                      { 
                        key: 'Title',
                        value: newValue == null || newValue == "" ? undefined : `startswith(Title,'${newValue}')`,
                        // type: 'text'
                      }
                    ]
                  }));
                }, 1000)
              }
              onChange={debounce((ev, newValue) => {
                    console.log("onChange Text", newValue); 
                    this.setState((prevState => ({ 
                      // searchText: newValue ==null || newValue == ""?undefined:newValue 
                      employees: [],
                      filterOptions: [
                        ...(prevState.filterOptions || []),
                        { 
                          key: 'Title',
                          value: newValue == null || newValue == "" ? undefined : `startswith(Title,'${newValue}')`,
                          // type: 'text'
                        }
                      ]
                    })
                  ))
                }, 1000)
              }
              styles={{ root: { width: 200 } }}
              ariaLabel="Search Employee"
              iconProps={{ iconName: 'Search' }}
              clearButtonProps={{ ariaLabel: 'Clear search' }}
              // onClear={(ev) => {
              //   console.log("onClear Text", ev); this.setState({ searchText: undefined })
              // }}
              underlined={true}
            />
          </div>
        </div>
        <div ref={this.employeeListRef} style={{ height: '500px', overflow: 'auto' }}>
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
