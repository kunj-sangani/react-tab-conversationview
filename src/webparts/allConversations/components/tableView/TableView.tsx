import { IColumn, SelectionMode, DetailsList } from 'office-ui-fabric-react';
import * as React from 'react';
import { ChatMessage as GraphChatMessage, ChatMessageAttachment } from "@microsoft/microsoft-graph-types";
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
import ReadMoreAndLess from 'react-read-more-less';
import { ViewType } from '@microsoft/mgt-spfx';
import { Person } from '@microsoft/mgt-react/dist/es6/spfx';
import { IGraphService } from '../../../../service/GraphService';
import styles from '.././AllConversations.module.scss';
import { ExpandIcon, ReplyIcon } from '@fluentui/react-northstar';

export interface ITableViewProps {
    expandedMessageId:string;
    graphService : IGraphService;
    getMessageReply: (messageId: string) => Promise<void>;
    dtaaFilteredMessage: GraphChatMessage[];
    dtaaFilteredMessageData: GraphChatMessage[];
}

export const TableView: React.FunctionComponent<ITableViewProps> = (props: React.PropsWithChildren<ITableViewProps>) => {

/**
 * converts message body to string
 * @param html input body of message
 * @returns string formated text
 */
 const convertToPlain = (html:string):string => {
    const tempDivElement = document.createElement("div");
    tempDivElement.innerHTML = html;
    return tempDivElement.textContent || tempDivElement.innerText || "";
}

const columns: IColumn[] = [
    {
        key: 'column3',
        name: 'createdBy',
        minWidth: 200,
        maxWidth: 200,
        isResizable: true,
        isCollapsible: true,
        onRender: (chatMessage: GraphChatMessage) => (
            <Person userId={chatMessage.from.user.id} view={ViewType.oneline} />
        )
    },
    {
        key: 'column1',
        name: 'Message',
        minWidth: 300,
        maxWidth: 300,
        isResizable: true,
        isCollapsible: true,
        onRender: (chatMessage: GraphChatMessage) => (
        <ReadMoreAndLess
            className="read-more-content"
            charLimit={50}
            readMoreText="Read more"
            readLessText=" Read less"
            >
            {convertToPlain(chatMessage.body.content)}
        </ReadMoreAndLess>
        )
    },
    {
        key: 'column2',
        name: 'created Date Time',
        fieldName: 'createdDateTime',
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        isCollapsible: true
    },
    {
        key: 'column4',
        name: 'attachment',
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        isCollapsible: true,
        onRender: (chatMessage: GraphChatMessage) => (
            <div className={styles.docWrapper}>
                {
                    chatMessage.attachments && chatMessage.attachments.map((at:ChatMessageAttachment,
                    attachmentIndex:number) => 
                        <a key ={`attachment${attachmentIndex}`} style={{cursor:'pointer'}} onClick={() => 
                        props.graphService.spcontext.sdks.microsoftTeams.teamsJs.app.openLink(at.contentUrl)} >{at.name}</a>
                    )
                }
            </div>
        )
    },
    {
        key: 'column5',
        name: 'Go to Message',
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        isCollapsible: true,
        onRender: (chatMessage: GraphChatMessage) => (
            <ExpandIcon onClick={()=> props.graphService.spcontext
                .sdks.microsoftTeams.teamsJs.app.openLink(chatMessage.webUrl)} />
        )
    },
    {
        key: 'column5',
        name: 'View Replies',
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        isCollapsible: true,
        onRender: (chatMessage: GraphChatMessage) => (
            !props.expandedMessageId && <ReplyIcon onClick={()=>props.getMessageReply(chatMessage.id)} />
        )
    }
];

  return (
    <>
    {!!props.expandedMessageId && props.dtaaFilteredMessageData.length > 0 && <DetailsList 
      items={props.dtaaFilteredMessageData} 
      columns={columns} selectionMode={SelectionMode.none} />}
      {props.dtaaFilteredMessage.length > 0 && 
      <DetailsList 
      items={props.dtaaFilteredMessage.filter(t=>t.from && t.from.user)} 
      columns={columns} selectionMode={SelectionMode.none} />}
    </>
  );
};