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
    textDecoration : 'line-through'
  },
  ToDoListItemNotCompleted: {

  }
}

/**
 * The 'PersonalTab' component renders the main tab content
 * of your app.
 */
class Tab extends React.Component<any> {

  constructor(props: any) {
    super(props)
    this.state = {
      userInfo: {},
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

  addNewTask = () => {
    this.setStage({addNewTask: Date.now() });
    console.log('NOW:', this.state.addNewTask);
  }

  render() {

    var todoItems = [
      { "name": "Add a task", isCompleted: false, isNew: true },
      { "name": "Task 1", isCompleted: false },
      { "name": "Task 2", isCompleted: false },
      { "name": "Task 3", isCompleted: true }
    ];

    // TODO: fix context
    // let context = this;

    return (
      <div className="Tab">
        <div className="Title">ToDo App</div>
        <div className="Subtitle">Hello, {this.state.userInfo.userName}</div>
        <div className="ToDoList">
          {
            todoItems.map(function (todoItem, index) {
              if (todoItem.isNew) {
                return (
                  <li className="ToDoListItem" key={index}>
                    <Input 
                      placeholder={todoItem.name} 
                      clearable 
                      icon={<AddIcon />} 
                      iconPosition="start" 
                      // fluid
                      styles={styles.ToDoListItemNew}
                      input={{
                        styles:{
                          background: 'transparent',
                        }
                      }}>
                      </Input>
                      <Button 
                        icon={<NotesIcon />} 
                        text
                        iconOnly
                        onClick={this.addNewTask}
                      >
                      </Button>
                  </li>
                )
              } else {
                return (
                  <li className="ToDoListItem" key={index}>
                    <Checkbox checked={todoItem.isCompleted} />
                    <Text 
                      styles={todoItem.isCompleted ? styles.ToDoListItemCompleted : styles.ToDoListItemNotCompleted}>{todoItem.name}
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
