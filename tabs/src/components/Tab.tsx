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
  toDoItems: any,
  newToDoItem: any,
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
      toDoItems: [
        { "name": "Add a task", isCompleted: false, isNew: true },
        { "name": "Task 1", isCompleted: false },
        { "name": "Task 2", isCompleted: false },
        { "name": "Task 3", isCompleted: true }
      ],
      newToDoItem: `test ${Date.now()}`,
    }
  }

  //React lifecycle method that gets called once a component has finished mounting
  //Learn more: https://reactjs.org/docs/react-component.html#componentdidmount
  async componentDidMount() {
    // Next steps: Error handling using the error object
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

  addNewTask(toDoItem: any, event: any) {
    if (toDoItem.name === "")
      return;

    console.log(toDoItem.name);
    this.state.toDoItems.splice(1, 0, { "name": toDoItem.name, isCompleted: false });
    this.state.newToDoItem = "";
    this.setState({ ...this.state });
  }

  handleNewToDoItemChange(toDoItem: any, event: any) {
    this.setState({ newToDoItem: event.target.value });
  }

  handleToDoItemCompletionChange(toDoItem: any, event: any) {
    toDoItem.isCompleted = event.target.value;
    this.setState({ ...this.state });
  }

  render() {
    return (
      <div className="Tab">
        <div className="Title">ToDo App</div>
        <div className="Subtitle">Hello, {this.state.userInfo.userName}</div>
        <div className="ToDoList">
          {
            this.state.toDoItems.map((toDoItem: any, index: any) => {
              if (toDoItem.isNew) {
                return (
                  <li className="ToDoListItem" key={index}>
                    <Input
                      placeholder={toDoItem.name}
                      clearable
                      icon={<AddIcon />}
                      iconPosition="start"
                      value={this.state.newToDoItem}
                      onChange={this.handleNewToDoItemChange.bind(this, toDoItem)}
                      // fluid
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
                      onClick={this.addNewTask.bind(this, toDoItem)}
                    />
                  </li>
                )
              } else {
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
              }
            })
          }
        </div>
      </div>
    );
  }
}

export default Tab;
