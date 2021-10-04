import * as React from 'react';
import styles from './GlossaryWp.module.scss';
import { IGlossaryWpProps } from './IGlossaryWpProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Accordion from 'react-bootstrap/Accordion';
import Card from 'react-bootstrap/Card';
import Pagination from 'react-bootstrap/Pagination';
import OverlayTrigger from 'react-bootstrap/OverlayTrigger';
import Tooltip from 'react-bootstrap/Tooltip';
import { SPHttpClient } from '@microsoft/sp-http';

import 'bootstrap/dist/css/bootstrap.min.css';

export interface IGlossaryViewState {
  items: any[];
  filtered: any[];
  def: any[];
  currentPage: string;
}

export default class GlossaryWp extends React.Component<IGlossaryWpProps, IGlossaryViewState, {}> {
  constructor(props: IGlossaryWpProps, state: IGlossaryViewState) {
    super(props);
    this.state = {
      items: [],
      def: [],
      filtered: [],
      currentPage: "A"
    };
  }

  public componentDidMount() {
    this._loadListItem();
  }

  public componentDidUpdate() {
    // this._loadListItem();
  }

  private genCharArray(charA, charZ) {
    var a = [], i = charA.charCodeAt(0), j = charZ.charCodeAt(0);
    for (; i <= j; ++i) {
        a.push(
          <Pagination.Item key={String.fromCharCode(i)} active={String.fromCharCode(i) === this.state.currentPage} onClick={(e) => this.handlePagination(e)}> 
            {String.fromCharCode(i)}
          </Pagination.Item>);
    }
    return a;
  }
  private items = [{id:1, title: "Arjon", content: "Hello"},{id:2, title: "Barjon", content: "Hello there"}];
  private handlePagination(e) {
    // this._loadListItem();
    let filtered = this.state.items.filter(function (items) {
      return items.title.charAt(0) === e.target.text;
    })
    this.setState({
      filtered: filtered ? filtered : [],
      currentPage: e.target.text
    });
  }

  public _loadListItem() { 
    const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/items`;
    const restApi1 = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.props.definitions}')/items`;
    this.props.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); })
      .then(items => {
        this.props.context.spHttpClient.get(restApi1, SPHttpClient.configurations.v1)
        .then(resp1 => { return resp1.json(); })
        .then(items1 => {
          let tempDef = [];
          items1.value.forEach((item) => {
            tempDef.push({title: item.Title, definition: item.DefinitionDescription})
          });
          this.setState({
            def: tempDef ? tempDef : []
          });
          console.log(this.state.def)
        });
        let tempIt = [];
        items.value.forEach((item, index) => {
          tempIt.push({id: index, title: item.Title, content: item.GlossaryDesc})
        });
        this.setState({
          items: tempIt ? tempIt : [],
          filtered: tempIt ? tempIt.filter(function (items) { return items.title.charAt(0) === 'A';}) : []
        });
      });
  }

  private constr(content, defs) {
    console.log(defs)
    var words = content.split(" ");
    var finalCont = "";
    console.log(words, defs)
    words.forEach((word, i) => {
      defs.forEach(def => {
          if (def.title === word) {
            console.log(words[i])
            words[i] = <OverlayTrigger overlay={<Tooltip>{def.definition}</Tooltip>}>
                  <span style={{textDecorationLine: 'underline', textDecorationStyle: 'dashed'}}>
                    {def.title}
                  </span>
                </OverlayTrigger>;
          }
      })
      if (typeof(words[i]) === "string") {
        words[i] = words[i] + " ";
        if (words[i - 1]) {
          if (typeof(words[i - 1]) !== "string") {
            words[i] = " " + words[i];
          }
        }
      } else {
        words[i] = words[i]
      }
    })
    return words
  }

  public render(): React.ReactElement<IGlossaryWpProps> {
    return (
      <div>
        <h1>Glossary</h1>
        <Pagination style={{justifyContent: 'left'}} size="sm">{this.genCharArray("A", "Z")}</Pagination>
        {this.state.filtered.map((item) => {
        return <Accordion key={(item.id === 0 ? 1 : item.id)} defaultActiveKey={this.state.items.length + 1}>
          <Card>
            <Accordion.Toggle as={Card.Header} eventKey={(item.id === 0 ? 1 : item.id)}>
              {item.title}
            </Accordion.Toggle>
            <Accordion.Collapse eventKey={(item.id === 0 ? 1 : item.id)}>
              <Card.Body>
              {this.constr(item.content, this.state.def)}</Card.Body>
            </Accordion.Collapse>
          </Card>
        </Accordion>})}
      </div>
    );
  }
}
