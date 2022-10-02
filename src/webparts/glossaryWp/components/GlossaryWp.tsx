import * as React from 'react';
import { IGlossaryWpProps } from './IGlossaryWpProps';
import Accordion from 'react-bootstrap/Accordion';
import Card from 'react-bootstrap/Card';
import Pagination from 'react-bootstrap/Pagination';
import { SPHttpClient } from '@microsoft/sp-http';
import Form from 'react-bootstrap/Form';
import ReactTooltip from 'react-tooltip';

import 'bootstrap/dist/css/bootstrap.min.css';

export interface IGlossaryViewState {
  items: any[];
  filtered: any[];
  def: any[];
  currentPage: string;
  contentHeaders: any[];
}

export default class GlossaryWp extends React.Component<IGlossaryWpProps, IGlossaryViewState, {}> {
  constructor(props: IGlossaryWpProps, state: IGlossaryViewState) {
    super(props);
    this.state = {
      items: [],
      def: [],
      filtered: [],
      currentPage: "A",
      contentHeaders: []
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

  private handlePagination(e) {
    var headers = this.state.contentHeaders;
    let filtered = this.state.items.filter(function (items) {
      return items[headers[0]].charAt(0) === e.target.text;
    })
    this.setState({
      filtered: filtered ? filtered : [],
      currentPage: e.target.text
    });
  }

  public _loadListItem () { 
    const glossaryItemsApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/items`;
    const definitApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.props.definitions}')/items`;
    const glossaryFieldsApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/fields?$filter=ReadOnlyField eq false and Hidden eq false`;

    const getResponse = async (api) => {
      return this.props.context.spHttpClient.get(api, SPHttpClient.configurations.v1)
      .then(response => {
        if (response.ok) {
          return response.json();
        }
      });
    };

    // get glossary fields
    const glossaryFields = getResponse(glossaryFieldsApi);
    // get glossary items
    const glossaryItems = getResponse(glossaryItemsApi);
    glossaryFields.then(resp => {
      var tmp = [];
      resp.value.forEach(val => {
        if (val.Title !== "Attachments" && val.Title !== "Content Type") {
          tmp.push(val.Title);
        }
      });
      this.setState({
        contentHeaders: tmp.length > 0 ? tmp : []
      });

      glossaryItems.then(resp => {
        var tmp1 = [];
        resp.value.forEach((item, index) => {
          var obj = {id: index + 1};
          this.state.contentHeaders.forEach(ch => {
            obj[ch] = item[ch] ? item[ch] : "";
          });
          tmp1.push(obj)
        });
        this.setState({
          items: tmp1.length > 0 ? tmp1 : [],
          filtered: tmp1.length > 0 ? tmp1.filter(function (items) { return items[tmp[0]].charAt(0) === 'A';}) : []
        });
      });
    });

    // get definitions
    if (!this.props.definitions) {
      return;
    }
    const definitions = getResponse(definitApi);
    definitions.then(resp => {
      var tmp = [];
      resp.value.forEach((item, index) => {
        tmp.push({title: item.Title, definition: item.DefinitionDescription})
      });
      this.setState({
        def: tmp.length > 0 ? tmp : []
      });
    });
  }

  public handleSearchBar = (e) => {
    var keyword = e.target.value;
    var keywordL = keyword.length;
    var headers = this.state.contentHeaders;
    if (keyword != "" && keywordL >= 1) {
      let filtered = this.state.items.filter(function (items) {
        return items[headers[0]].toLowerCase().includes(keyword.toLowerCase());
      });
      filtered.sort(function(a, b) {
        return a[headers[0]].localeCompare(b[headers[0]])
      });
      this.setState({
        filtered: filtered.length !== 0 ? filtered : [],
        currentPage: ''
      });
    }
  }

  private constr(content, defs) {
    // definitions
    const addDefin = (cont) => {
      var words = cont.split(" ");
      words.forEach((word, i) => {
        defs.forEach(def => {
            if (word.includes(def.title)) {
              words[i] = `<span data-tip=${def.definition} style="text-decoration: underline dotted">${word}</span>`;
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
      return words.join(" ");
    };

    var headrs = this.state.contentHeaders.filter(hd => hd !== this.state.contentHeaders[0]);
    var fullContent = "";
    headrs.forEach(header => {
      var titleH = `<h6>${header}</h6>`;
      var body = `<p>${addDefin(content[header])}</p>`;
      fullContent += (titleH + body);
    });

    return <div> <div dangerouslySetInnerHTML={{ __html: fullContent }}></div>
        <ReactTooltip effect="solid" />
      </div>
  }

  public render(): React.ReactElement<IGlossaryWpProps> {
    return (
      <div>
        <h2 className='text-center'>Glossary</h2>
        <Form className="d-flex" style={{padding: '10px'}}>
          <Form.Control
            type="search"
            placeholder="Search"
            className="me-2"
            aria-label="Search"
            onChange={this.handleSearchBar}
          />
        </Form>
        <Pagination style={{justifyContent: 'center'}} size="sm">{this.genCharArray("A", "Z")}</Pagination>
        {(this.state.filtered.length != 0) && this.state.filtered.map((item) => {
        return <Accordion key={item.id}>
          <Card style={{ width: '45rem' }}>
            <Accordion.Toggle as={Card.Header} eventKey={item.id}>
              {item[this.state.contentHeaders[0]]}
            </Accordion.Toggle>
            <Accordion.Collapse eventKey={item.id}>
              <Card.Body>
              {this.constr(item, this.state.def)}</Card.Body>
            </Accordion.Collapse>
          </Card>
        </Accordion>})}
      </div>
    );
  }
}
