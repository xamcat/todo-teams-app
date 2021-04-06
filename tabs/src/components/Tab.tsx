// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React, { Component, useState } from 'react';
import './App.css';
import './Tab.css';

import { MODS } from "mods-client";
import * as microsoftTeams from "@microsoft/teams-js";
import { Text, Input, Button, Checkbox, AddIcon, ToDoListIcon, EditIcon } from '@fluentui/react-northstar';
import SlidingPanel from 'react-sliding-side-panel';

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
  toDoItems: any,
  isDoItemDetailsOpened: boolean,
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
      toDoItems: [
        { name: "Task 1", isCompleted: false },
        { name: "Task 2", isCompleted: false },
        { name: "Task 3", isCompleted: true }
      ],
      isDoItemDetailsOpened: true,
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
    this.setState({ ...this.state });
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
    // console.log(`item ${toDoItem.name} has been selected`);
    this.state.isDoItemDetailsOpened = !this.state.isDoItemDetailsOpened;
    this.setState({ isDoItemDetailsOpened: this.state.isDoItemDetailsOpened });
  }

  render() {
    return (
      <div className="Tab">
        <div className="Title">ToDo App</div>
        <div className="Subtitle">Hello, {this.state.userInfo.userName}</div>
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
              icon={<EditIcon />}
              text
              iconOnly
              onClick={this.addNewTask.bind(this, this.state.toDoItemNew)}
            />
            <Button
              icon={<ToDoListIcon />}
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
        <SlidingPanel
            type={'right'}
            isOpen={this.state.isDoItemDetailsOpened}
            size={30}
          >
            <div>
              <div>My Panel Content</div>
              <button onClick={() => this.setState({ isDoItemDetailsOpened: false })}>close</button>
            </div>
          </SlidingPanel>
      </div>
    );
  }
}

export default Tab;
