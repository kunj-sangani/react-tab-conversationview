import * as React from 'react';
import { Text,  Button, Flex} from '@fluentui/react-northstar';
import { useState } from 'react';
import { DatePicker, DayOfWeek, SearchBox, Checkbox } from 'office-ui-fabric-react';
import { PeoplePicker } from '@microsoft/mgt-react/dist/es6/spfx';
import { ChatMessage as GraphChatMessage, User } from "@microsoft/microsoft-graph-types";

export interface IFilterProps {
    allChatMessage:GraphChatMessage[];
    onFilter:(searchText: string, fromDate: Date, toDate: Date, onlyAttachment:boolean,
    fromUser:Array<User>, mentionedUser:Array<User>) => void;
    isOpen: boolean;
    flipFilterPanel: (isOpen: boolean) => void;
}


export const Filter: React.FunctionComponent<IFilterProps> = (props: React.PropsWithChildren<IFilterProps>) => {
    const [searchText, setSearchText] = useState('');
    const [fromUser, setFromUser] = useState<Array<User>>([]);
    const [mentionedUser, setMentionedUser] = useState<Array<User>>([]);
    const [fromDate, setFromDate] = useState<Date>();
    const [toDate, setToDate] = useState<Date>();
    const [onlyAttachment, setOnlyAttachment] = useState(false);

    /**
     * Function on click of submit
     */
    const onSubmit = ():void => {
        props.onFilter(searchText, fromDate, toDate, onlyAttachment, fromUser, mentionedUser);
        props.flipFilterPanel(false);
    }

    /**
     * Function on click of reset
     */
    const onReset = ():void => {
        setSearchText('');
        setFromUser([]);
        setMentionedUser([]);
        setFromDate(null);
        setToDate(null);
        setOnlyAttachment(false);
        props.onFilter('', null, null,false,[],[]);
        props.flipFilterPanel(false);
    }

    /**
     * display filter and reset button in panel
     * @returns JSX element for footer in the panel
     */
    const onRenderFooterContent = ():JSX.Element => {
        return (
            <Flex gap="gap.smaller" vAlign='end'>
                <Button content="Apply" primary onClick={onSubmit} />
                <Button content="Reset" onClick={onReset}/>
          </Flex>
          )
    }
    
  return (
    <Flex gap="gap.small">
        {/* <Panel
            headerText="Filter panel"
            isOpen={props.isOpen}
            onDismiss={()=>props.flipFilterPanel(false)}
            closeButtonAriaLabel="Close"
            onRenderFooterContent={onRenderFooterContent}
            isFooterAtBottom={true}
        > */}
            <Flex column gap='gap.medium'>
                <Flex gap='gap.medium'>
                    <Flex column style={{minWidth:350}} >
                        <Text content="Body Serach" />
                        <SearchBox placeholder="Search" value={searchText} onChange={(_,newValue)=>setSearchText(newValue)} />
                    </Flex>
                    <Flex column style={{minWidth:300}}>
                        <Text content="From" />
                        <PeoplePicker userIds={props.allChatMessage.filter(t=>t.from && t.from.user).map((t)=>t.from.user.id)}
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        selectionChanged={(e:any)=>setFromUser(e.target.selectedPeople)} selectedPeople={fromUser} />
                    </Flex>
                    <Flex column style={{minWidth:300}}>
                        <Text content="Mentions" />
                        <PeoplePicker  userIds={[].concat(...props.allChatMessage.filter(t=>t.mentions.length > 0).map((t)=>t.mentions.filter(m=>m.mentioned.user).map((m)=>m.mentioned.user.id)))}
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        selectionChanged={(e:any)=>setMentionedUser(e.target.selectedPeople)} selectedPeople={mentionedUser} />
                    </Flex>
                </Flex>
                <Flex gap='gap.medium' style={{marginBottom:10}}>
                    <Flex column style={{minWidth:300}}>
                        <Text content="From Date" />
                        <DatePicker firstDayOfWeek={DayOfWeek.Monday} placeholder="Select a From date..."
                        onSelectDate={(date)=>setFromDate(date)} initialPickerDate={fromDate} value={fromDate} />
                    </Flex>
                    <Flex column style={{minWidth:300}}>
                        <Text content="To Date" />
                        <DatePicker firstDayOfWeek={DayOfWeek.Monday} placeholder="Select a To date..." 
                        onSelectDate={(date)=>setToDate(date)} initialPickerDate={toDate} value={toDate} />
                    </Flex>
                    <Flex column vAlign='end' style={{minWidth:135}}>
                        <Checkbox label="Attachments" checked={onlyAttachment} onChange={(_,checked)=>setOnlyAttachment(checked)}  />
                    </Flex>
                    {onRenderFooterContent()}
                </Flex>
            </Flex>
        {/* </Panel> */}
      </Flex>
  );
};