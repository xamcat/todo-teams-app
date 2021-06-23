import React from "react";
import { TeamsUserCredential, createMicrosoftGraphClient, getResourceConfiguration, ResourceType } from "@microsoft/teamsfx";
import * as microsoftTeams from "@microsoft/teams-js";
import { Text, Input, Button, Image, Checkbox, Datepicker, AddIcon, LockIcon, ParticipantAddIcon, TrashCanIcon, SendIcon, RaiseHandIcon, ArrowRightIcon, FilesImageIcon, NotesIcon, LightningIcon } from '@fluentui/react-northstar';
import { v4 as uuid } from "uuid";
import * as axios from "axios";
import './App.css';
import './Tab.css';

const styles = {
  ToDoListItemNew: {
  },
  ToDoListItemCompleted: {
    color: 'gray',
    textDecoration: 'line-through'
  },
  ToDoListItemNotCompleted: {
    color: 'black',
    textDecoration: 'none'
  }
}

const localStorageKey = "ToDoItemsState.v1.0.4";
const graphApiScopes = ['User.Read', 'User.ReadBasic.All'];
const azureFunctionName = process.env.REACT_APP_FUNC_NAME || "myFunc";

interface TabProps {
}

interface TabState {
  userInfo: any,
  toDoItemNew: any,
  toDoItemDetails: any,
  toDoItems: any,
}

class Tab extends React.Component<TabProps, TabState> {

  private filePickerRef: any;

  constructor(props: TabProps) {
    super(props)
    this.filePickerRef = React.createRef();
    this.state = {
      userInfo: {},
      toDoItemNew: {
        id: "",
        name: "",
        isCompleted: false,
        notes: "",
        dueDate: null,
        createdDate: null,
        people: null,
        attachments: [],
        label: "Add a task"
      },
      toDoItemDetails: null,
      toDoItems: [],
    }
  }

  async componentDidMount() {
    this.loadStateFromStorage();
    await this.initTeamsFx();
    microsoftTeams.initialize();
  }

  async initTeamsFx() {
    const credential = new TeamsUserCredential();
    var userInfo = await credential.getUserInfo();
    this.setState({
      userInfo: userInfo
    });
  }

  async authorizeTeamsFx() {
    const credential = new TeamsUserCredential();
    await credential.login(graphApiScopes);
  }

  async createGraphClient() {
    const credential = new TeamsUserCredential();
    const graph = createMicrosoftGraphClient(credential, graphApiScopes);
    return graph;
  }

  async searchUsers(filter: string) {
    if (!filter || filter === '')
      return;

    try {
      const graph = await this.createGraphClient();
      var graphQuery = `/users?$filter=startswith(displayName,'${filter}')`;
      var searchUsers = await graph.api(graphQuery).get();
      if (searchUsers && searchUsers.value && searchUsers.value.length > 0) {
        const bestMatch = searchUsers.value[0];
        console.log(`Found ${searchUsers.value.length} users, assigning ${bestMatch.userPrincipalName}...`);
        this.state.toDoItemDetails.people = bestMatch.userPrincipalName;
        this.setState({ toDoItemDetails: this.state.toDoItemDetails });
      }
    }
    catch (err) {
      console.log(err);
      await this.authorizeTeamsFx();
    }
  }

  async searchManager() {
    try {
      const graph = await this.createGraphClient();
      var bestMatch = await graph.api(`/me/manager`).get();
      if (bestMatch) {
        console.log(`Found manager, assigning ${bestMatch.userPrincipalName}...`);
        this.state.toDoItemDetails.people = bestMatch.userPrincipalName;
        this.setState({ toDoItemDetails: this.state.toDoItemDetails });
      }
    }
    catch (err) {
      console.log(err);
      await this.authorizeTeamsFx();
    }
  }

  loadStateFromStorage() {
    const localState = localStorage.getItem(localStorageKey);
    if (localState) {
      const restoredState = JSON.parse(localState);
      if (restoredState.toDoItemDetails !== null) {
        restoredState.toDoItemDetails = restoredState.toDoItems.find((i: any) => i.id === restoredState.toDoItemDetails.id);
      }
      this.setState({ ...restoredState });
    } else {
      const defaultState = [
        { id: uuid(), name: "Task 1", isCompleted: false, notes: "Notes for task 1", dueDate: null, createdDate: Date.now(), people: null, attachments: [] },
        { id: uuid(), name: "Task 2", isCompleted: false, notes: "Notes for task 2", dueDate: null, createdDate: Date.now(), people: null, attachments: [] },
        { id: uuid(), name: "Task 3", isCompleted: true, notes: "Notes for task 3", dueDate: null, createdDate: Date.now(), people: null, attachments: [] }
      ]
      this.setState({ toDoItems: defaultState }, () => this.saveStateToStorage());
    }
  }

  saveStateToStorage() {
    var localState = JSON.stringify(this.state);
    localStorage.setItem(localStorageKey, localState);
  }

  selectMedia() {
    let imageProp: microsoftTeams.media.ImageProps = {
      sources: [microsoftTeams.media.Source.Gallery],
      startMode: microsoftTeams.media.CameraStartMode.Photo,
      ink: false,
      cameraSwitcher: false,
      textSticker: false,
      enableFilter: true,
    };

    let mediaInput: microsoftTeams.media.MediaInputs = {
      mediaType: microsoftTeams.media.MediaType.Image,
      maxMediaCount: 10,
      imageProps: imageProp
    };

    // requests for access but then crashes the browser
    // navigator.mediaDevices.getUserMedia({ audio: true, video: true });
    microsoftTeams.media.selectMedia(mediaInput, (error: microsoftTeams.SdkError, files: microsoftTeams.media.File[]) => {
      if (error) {
        console.log(error);
        this.filePickerRef.current.click();
      } else {
        console.log(`Do work with images: ${files.length}`);
      }
    });
  }

  async callAzureFunction() {
    try {
      const credential = new TeamsUserCredential();
      const accessToken = await credential.getToken("");
      const apiConfig = getResourceConfiguration(ResourceType.API);
      const response = await axios.default.get(apiConfig.endpoint + "/api/" + azureFunctionName, {
        headers: {
          authorization: "Bearer " + accessToken?.token || "",
        },
      });

      console.log(response);
      window.alert(`Response from the Azure Function [${azureFunctionName}]:\n${response?.data?.userInfoMessage}`);
    } catch (err) {
      console.log(err);
      await this.authorizeTeamsFx();
    }
  }

  addNewTask(toDoItem: any) {
    if (toDoItem.name === "") {
      return;
    }

    this.state.toDoItems.splice(0, 0, {
      id: uuid(),
      name: toDoItem.name,
      isCompleted: false,
      notes: "",
      dueDate: null,
      createdDate: Date.now(),
      people: null,
      attachments: []
    });
    toDoItem.name = "";
    this.setState({ toDoItemNew: this.state.toDoItemNew }, () => this.saveStateToStorage());
  }

  clearAllTasks() {
    if (!window.confirm("Do you want to delete all tasks?"))
      return;

    this.setState({
      toDoItems: [],
      toDoItemDetails: null,
    }, () => this.saveStateToStorage());
  }

  removeToDoTask(toDoItem: any) {
    if (!toDoItem) {
      return;
    }

    if (this.state.toDoItemDetails === toDoItem) {
      this.setState({ toDoItemDetails: null }, () => this.saveStateToStorage());
    }

    const index = this.state.toDoItems.indexOf(toDoItem);
    if (index > -1) {
      this.state.toDoItems.splice(index, 1);
      this.saveStateToStorage();
    }
  }

  handleToDoItemDetailsNameChange(toDoItem: any, event: any) {
    toDoItem.name = event.target.value;
    this.setState({ toDoItemDetails: this.state.toDoItemDetails });
  }

  handleToDoItemDetailsNotesChange(toDoItem: any, event: any) {
    toDoItem.notes = event.target.value;
    this.setState({ toDoItemDetails: this.state.toDoItemDetails });
  }

  handleToDoItemDetailsDateChange(toDoItem: any, event: any, args: any) {
    toDoItem.dueDate = args.itemProps.value.originalDate.getTime();
    this.setState({ toDoItemDetails: this.state.toDoItemDetails });
  }

  async handleToDoItemDetailsPeopleChange(toDoItem: any, event: any) {
    toDoItem.people = event.target.value;
    this.setState({ toDoItemDetails: this.state.toDoItemDetails });
  }

  async handleToDoItemDetailsPeopleKeyPress(toDoItem: any, event: any) {
    switch (event.key) {
      case "Escape":
        toDoItem.people = "";
        this.setState({ toDoItemDetails: this.state.toDoItemDetails });
        break;
      case "Enter":
        await this.searchUsers(event.target.value);
        break;
    }
  }

  async handleToDoItemDetailsPeopleBlur(toDoItem: any, event: any) {
    await this.searchUsers(event.target.value);
  }

  handleToDoItemDetailsAttachmentsChange(toDoItem: any, event: any) {
    if (event.target.value === null)
      return;

    console.log(`New attachment: ${event.target.value}`);

    const file = this.filePickerRef.current.files[0];
    console.log(`The file: ${file.name}`);

    var reader = new FileReader();
    reader.onloadend = () => {
      const defaultImage = "https://www.pngarts.com/files/2/Upload-Free-PNG-Image.png";
      this.state.toDoItemDetails.attachments.push({
        name: uuid(),
        previewSource: (file.type === "image/png" || file.type === "image/jpeg") ? [reader.result] : defaultImage,
        isUploaded: false,
      });
      this.setState({ toDoItemDetails: this.state.toDoItemDetails }, () => this.saveStateToStorage());
      this.filePickerRef.current.value = null;
    };

    reader.readAsDataURL(file);
  }

  handleNewToDoItemChange(toDoItem: any, event: any) {
    toDoItem.name = event.target.value;
    this.setState({ toDoItemNew: this.state.toDoItemNew });
  }

  handleNewToDoItemKeyPress(toDoItem: any, event: any) {
    switch (event.key) {
      case "Escape":
        toDoItem.name = "";
        this.setState({ toDoItemNew: this.state.toDoItemNew });
        break;
      case "Enter":
        this.addNewTask(toDoItem);
        break;
    }
  }

  handleNewToDoItemBlur(toDoItem: any, event: any) {
    this.saveStateToStorage();
  }

  handleToDoItemCompletionChange(toDoItem: any, event: any) {
    toDoItem.isCompleted = !toDoItem.isCompleted;
    this.setState({ ...this.state }, () => this.saveStateToStorage());
  }

  handleToDoItemSelected(toDoItem: any, event: any) {
    this.setState({ toDoItemDetails: toDoItem }, () => this.saveStateToStorage());
  }

  render() {
    return (
      <div className="Tab">
        <div className="Title">ToDo App</div>
        <div className="Subtitle">Hello, {this.state.userInfo.displayName}</div>
        <div className="AddNewTask">
          <Input
            placeholder={this.state.toDoItemNew.label}
            clearable
            icon={<AddIcon />}
            iconPosition="start"
            value={this.state.toDoItemNew.name}
            onKeyDown={this.handleNewToDoItemKeyPress.bind(this, this.state.toDoItemNew)}
            onChange={this.handleNewToDoItemChange.bind(this, this.state.toDoItemNew)}
            onBlur={this.handleNewToDoItemBlur.bind(this, this.state.toDoItemNew)}
            styles={styles.ToDoListItemNew}
            input={{
              styles: {
                background: 'transparent',
              }
            }}>
          </Input>
          <Button
            icon={<SendIcon />}
            text
            iconOnly
            title="Add new task"
            onClick={this.addNewTask.bind(this, this.state.toDoItemNew)}
          />
          <Button
            icon={<TrashCanIcon />}
            text
            iconOnly
            title="Delete all tasks"
            onClick={this.clearAllTasks.bind(this)}
          />
          <Button
            icon={<LockIcon />}
            text
            iconOnly
            title="Authorize TeamsFx"
            onClick={this.authorizeTeamsFx.bind(this)}
          />
          <Button
            icon={<LightningIcon />}
            text
            iconOnly
            title="Call Azure Function"
            onClick={this.callAzureFunction.bind(this)}
          />
        </div>
        <div className="FlexContainer">
          <div className="FlexItemMain">
            <div className="ToDoList">
              {
                this.state.toDoItems.map((toDoItem: any, index: any) => {
                  return (
                    <li className={toDoItem.id !== this.state.toDoItemDetails?.id ? "ToDoListItem" : "ToDoListItemSelected"} key={index} onClick={this.handleToDoItemSelected.bind(this, toDoItem)}>
                      <Checkbox
                        className="ToDoListItemCheckbox"
                        checked={toDoItem.isCompleted}
                        onChange={this.handleToDoItemCompletionChange.bind(this, toDoItem)} />
                      <Text
                        styles={toDoItem.isCompleted ? styles.ToDoListItemCompleted : styles.ToDoListItemNotCompleted}>{toDoItem.name}
                      </Text>
                    </li>
                  )
                })
              }
            </div>
          </div>
          {this.state.toDoItemDetails !== null &&
            <div className="FlexItemDetails">
              <div className="FlexItemDetailsContent">
                <div className="FlexItemDetailsContentField">
                  <Checkbox
                    className="ToDoListItemCheckbox"
                    checked={this.state.toDoItemDetails.isCompleted}
                    toggle
                    onChange={this.handleToDoItemCompletionChange.bind(this, this.state.toDoItemDetails)}
                  />
                  <Input
                    value={this.state.toDoItemDetails.name}
                    onChange={this.handleToDoItemDetailsNameChange.bind(this, this.state.toDoItemDetails)}
                    input={{
                      styles: {
                        background: 'transparent',
                        width: 300,
                        textDecoration: this.state.toDoItemDetails.isCompleted ? 'line-through' : 'none'
                      }
                    }} />
                </div>
                <div className="FlexItemDetailsContentField">
                  <Input
                    clearable
                    icon={<NotesIcon />}
                    label="Notes"
                    labelPosition="inside"
                    value={this.state.toDoItemDetails.notes}
                    onChange={this.handleToDoItemDetailsNotesChange.bind(this, this.state.toDoItemDetails)}
                    input={{
                      styles: {
                        background: 'transparent',
                        width: 350
                      }
                    }} />
                </div>
                <div className="FlexItemDetailsContentField">
                  <Datepicker
                    allowManualInput={false}
                    daysToSelectInDayView={1}
                    selectedDate={(this.state.toDoItemDetails.dueDate ? new Date(this.state.toDoItemDetails.dueDate) : new Date())}
                    onDateChange={this.handleToDoItemDetailsDateChange.bind(this, this.state.toDoItemDetails)}
                    input={{
                      styles: {
                        background: 'transparent',
                        width: 315
                      }
                    }} />
                </div>
                <div className="FlexItemDetailsContentField">
                  <Input
                    clearable
                    icon={<ParticipantAddIcon />}
                    label="People"
                    labelPosition="inside"
                    value={this.state.toDoItemDetails.people}
                    onChange={this.handleToDoItemDetailsPeopleChange.bind(this, this.state.toDoItemDetails)}
                    onKeyDown={this.handleToDoItemDetailsPeopleKeyPress.bind(this, this.state.toDoItemDetails)}
                    input={{
                      styles: {
                        background: 'transparent',
                        width: 315
                      }
                    }}>
                  </Input>
                  <Button
                    icon={<RaiseHandIcon />}
                    text
                    iconOnly
                    title="Find my manager!"
                    onClick={this.searchManager.bind(this)}
                  />
                </div>
                <div className="FlexItemDetailsContentFieldAttachments">
                  <Button
                    icon={<FilesImageIcon />}
                    text
                    content="Attachments"
                    iconOnly
                    onClick={this.selectMedia.bind(this)} />
                  <Input
                    ref={this.filePickerRef}
                    type="file"
                    onChange={this.handleToDoItemDetailsAttachmentsChange.bind(this, this.state.toDoItemDetails)}
                    input={{
                      styles: {
                        background: 'transparent',
                        display: 'none'
                      }
                    }} />
                  <div className="FlexItemDetailsContentFieldAttachmentsList">
                    {this.state.toDoItemDetails.attachments?.map((attachment: any, index: any) => {
                      return (
                        <div className="FlexItemDetailsContentFieldAttachmentsPreview" key={index}>
                          <Image
                            className="FlexItemDetailsContentFieldAttachmentsPreviewImage"
                            src={attachment.previewSource}
                          />
                        </div>
                      )
                    })}
                  </div>

                </div>
              </div>
              <div className="FlexItemDetailsToolbar">
                <div className="FlexItemDetailsToolbarLeft">
                  <Button
                    icon={<ArrowRightIcon />}
                    text
                    iconOnly
                    onClick={this.handleToDoItemSelected.bind(this, null)} />
                </div>
                <div className="FlexItemDetailsToolbarMiddle">
                  <div>
                    Created {new Date(this.state.toDoItemDetails.createdDate).toLocaleString()}
                  </div>
                  <div className="ToDoListItemUuid">
                    {this.state.toDoItemDetails.id}
                  </div>
                </div>
                <div className="FlexItemDetailsToolbarRight">
                  <Button
                    icon={<TrashCanIcon />}
                    text
                    iconOnly
                    onClick={this.removeToDoTask.bind(this, this.state.toDoItemDetails)}
                  />
                </div>
              </div>
            </div>
          }
        </div>
      </div>
    );
  }
}

export default Tab;