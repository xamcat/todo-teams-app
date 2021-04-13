// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React, { Component, useState } from 'react';
import { MODS } from "mods-client";
import * as microsoftTeams from "@microsoft/teams-js";
import { Text, Input, Button, Checkbox, Datepicker, AddIcon, ToDoListIcon, MenuIcon, CloseIcon, TrashCanIcon, SendIcon, ArrowRightIcon, TeamCreateIcon, UserFriendsIcon, FilesImageIcon, CalendarIcon, NotesIcon } from '@fluentui/react-northstar';
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

interface TabProps {
}

interface TabState {
  userInfo: any,
  toDoItemNew: any,
  toDoItemDetails: any,
  toDoItems: any,
}

/**
 * The 'PersonalTab' component renders the main tab content
 * of your app.
 */
class Tab extends React.Component<TabProps, TabState> {

  constructor(props: TabProps) {
    super(props)
    this.state = {
      userInfo: {},
      toDoItemNew: { name: "", label: "Add a task" },
      toDoItemDetails: null,
      toDoItems: [
        { name: "Task 1", isCompleted: false },
        { name: "Task 2", isCompleted: false },
        { name: "Task 3", isCompleted: true }
      ],
    }
  }

  //React lifecycle method that gets called once a component has finished mounting
  //Learn more: https://reactjs.org/docs/react-component.html#componentdidmount
  async componentDidMount() {
    // Next steps: Error handling using the error object
    this.loadStateFromStorage();
    await this.initMODS();
    microsoftTeams.initialize();
  }

  async initMODS() {
    var modsEndpoint = process.env.REACT_APP_MODS_ENDPOINT;
    var startLoginPageUrl = process.env.REACT_APP_START_LOGIN_PAGE_URL;
    var functionEndpoint = process.env.REACT_APP_FUNC_ENDPOINT;
    await MODS.init(modsEndpoint!, startLoginPageUrl!, functionEndpoint);
    var userInfo = MODS.getUserInfo();
    this.setState({
      userInfo: userInfo
    });
  }

  loadStateFromStorage() {
    var localState = localStorage.getItem("ToDoItemsState");
    if (localState) {
      var restoredState = JSON.parse(localState);
      this.setState({ ...restoredState });
    }
  }

  saveStateToStorage() {
    var localState = JSON.stringify(this.state);
    localStorage.setItem("ToDoItemsState", localState);
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
      console.log(`Do work with images: ${error}`);
    });
  }

  addNewTask(toDoItem: any) {
    if (toDoItem.name === "") {
      return;
    }

    this.state.toDoItems.splice(0, 0, { name: toDoItem.name, isCompleted: false });
    this.state.toDoItemNew.name = "";
    this.setState({ ...this.state });
    this.saveStateToStorage();
  }

  clearAllTasks() {
    this.state.toDoItems = [];
    this.state.toDoItemDetails = null;
    this.setState({ ...this.state });
  }

  removeToDoTask(toDoItem: any) {
    if (!toDoItem) {
      return;
    }

    const index = this.state.toDoItems.indexOf(toDoItem);
    if (index > -1) {
      this.state.toDoItems.splice(index, 1);
    }

    if (this.state.toDoItemDetails === toDoItem) {
      this.state.toDoItemDetails = null;
    }

    this.setState({ ...this.state });
    this.saveStateToStorage();
  }

  handleNewToDoItemChange(toDoItem: any, event: any) {
    this.state.toDoItemNew.name = event.target.value;
    this.setState({ toDoItemNew: this.state.toDoItemNew });
  }

  handleNewToDoItemKeyPress(toDoItem: any, event: any) {
    switch (event.key) {
      case "Escape":
        this.state.toDoItemNew.name = "";
        this.setState({ toDoItemNew: this.state.toDoItemNew });
        break;
      case "Enter":
        this.addNewTask(toDoItem);
        break;
    }
  }

  handleNewToDoItemBlur(toDoItem: any, event: any) {
    console.log(event);
    this.saveStateToStorage();
  }

  handleToDoItemCompletionChange(toDoItem: any, event: any) {
    toDoItem.isCompleted = !toDoItem.isCompleted;
    this.setState({ ...this.state });
    this.saveStateToStorage();
  }

  handleToDoItemSelected(toDoItem: any, event: any) {
    if (!toDoItem) {
      return;
    }

    this.state.isDoItemDetailsOpened = (toDoItem && toDoItem != this.state.toDoItemDetails) ? true : false;
    this.state.toDoItemDetails = toDoItem;
    this.setState({ ...this.state });
  }

  handleToDoItemDeselected() {
    this.state.toDoItemDetails = null;
    this.setState({ ...this.state });
  }

  render() {
    return (
      <div className="Tab">
        <div className="Title">ToDo App</div>
        <div className="Subtitle">Hello, {this.state.userInfo.userName}</div>
        <div className="FlexContainer">
          <div className="FlexItemMain">
            <div className="ToDoList">
              <li className="ToDoListItem" key="-1">
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
                      // color: 'cornflowerblue',
                      background: 'transparent',
                    }
                  }}>
                </Input>
                <Button
                  icon={<SendIcon />}
                  text
                  iconOnly
                  onClick={this.addNewTask.bind(this, this.state.toDoItemNew)}
                />
                <Button
                  icon={<TrashCanIcon />}
                  text
                  iconOnly
                  onClick={this.clearAllTasks.bind(this)}
                />
              </li>
              {
                this.state.toDoItems.map((toDoItem: any, index: any) => {
                  return (
                    <li className="ToDoListItem" key={index} onClick={this.handleToDoItemSelected.bind(this, toDoItem)}>
                      <Checkbox
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
                    checked={this.state.toDoItemDetails.isCompleted}
                    onChange={this.handleToDoItemCompletionChange.bind(this, this.state.toDoItemDetails)} />
                  <Text
                    styles={this.state.toDoItemDetails.isCompleted ? styles.ToDoListItemCompleted : styles.ToDoListItemNotCompleted}>{this.state.toDoItemDetails.name}
                  </Text>
                </div>
                <div className="FlexItemDetailsContentField">
                  <Input
                    clearable
                    icon={<NotesIcon />}
                    label="Notes"
                    labelPosition="inside"
                    // value={this.state.toDoItemNew.name}
                    // onKeyDown={this.handleNewToDoItemKeyPress.bind(this, this.state.toDoItemNew)}
                    // onChange={this.handleNewToDoItemChange.bind(this, this.state.toDoItemNew)}
                    // onBlur={this.handleNewToDoItemBlur.bind(this, this.state.toDoItemNew)}
                    input={{
                      styles: {
                        background: 'transparent',
                        width: 350
                      }
                    }} />
                </div>
                <div className="FlexItemDetailsContentField">
                  {/* <CalendarIcon /> */}
                  <Datepicker
                    daysToSelectInDayView={1}
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
                    icon={<TeamCreateIcon />}
                    label="People"
                    labelPosition="inside"
                    // value={this.state.toDoItemNew.name}
                    // onKeyDown={this.handleNewToDoItemKeyPress.bind(this, this.state.toDoItemNew)}
                    // onChange={this.handleNewToDoItemChange.bind(this, this.state.toDoItemNew)}
                    // onBlur={this.handleNewToDoItemBlur.bind(this, this.state.toDoItemNew)}
                    input={{
                      styles: {
                        background: 'transparent',
                        width: 350
                      }
                    }}>
                  </Input>
                </div>
                <div className="FlexItemDetailsContentField">
                  <FilesImageIcon />
                  <Text>   [TBD] Attachments...</Text>
                </div>
              </div>
              <div className="FlexItemDetailsToolbar">
                <div className="FlexItemDetailsToolbarLeft">
                  <Button
                    icon={<ArrowRightIcon />}
                    text
                    iconOnly
                    onClick={this.handleToDoItemDeselected.bind(this)} />
                </div>
                <div className="FlexItemDetailsToolbarMiddle">
                  Created 2021-10-03 12:30 PM
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

// class ToDoItemDetails extends React.Component {
//   render() {
//     return (

//     );
//   }
// }

export default Tab;