// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import './Tab.css';
import { MODS } from "mods-client";
import { Text, Input, Button, Checkbox, AddIcon, NotesIcon } from '@fluentui/react-northstar';

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
    }
  }

  //React lifecycle method that gets called once a component has finished mounting
  //Learn more: https://reactjs.org/docs/react-component.html#componentdidmount
  async componentDidMount() {
    // Next steps: Error handling using the error object
    this.loadStateFromStorage();
    await this.initMODS();
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
    var localState = localStorage.getItem("ToDoItems");
    console.log(localState);

    // TODO: implement full state save/restore
    this.state.toDoItemNew.name = localState;
    this.setState({ toDoItemNew: this.state.toDoItemNew });
  }

  saveStateToStorage() {
    // TODO: implement full state save/restore
    localStorage.setItem("ToDoItems", this.state.toDoItemNew.name);
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

  handleNewToDoItemChange(toDoItem: any, event: any) {
    this.state.toDoItemNew.name = event.target.value;
    this.setState({ toDoItemNew: this.state.toDoItemNew });
    this.saveStateToStorage();
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

  handleToDoItemCompletionChange(toDoItem: any, event: any) {
    toDoItem.isCompleted = !toDoItem.isCompleted;
    this.setState({ ...this.state });
    this.saveStateToStorage();
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
              styles={styles.ToDoListItemNew}
              input={{
                styles: {
                  // color: 'cornflowerblue',
                  background: 'transparent',
                }
              }}>
            </Input>
            <Button
              icon={<NotesIcon />}
              text
              iconOnly
              onClick={this.addNewTask.bind(this, this.state.toDoItemNew)}
            />
          </li>
          {
            this.state.toDoItems.map((toDoItem: any, index: any) => {
              return (
                <li className="ToDoListItem" key={index}>
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
    );
  }
}

export default Tab;
