class App extends React.Component {
  render() {
    return (
      <div>
        <nav class="teal top-nav z-depth-0">
          <div class="container">
            <div class="nav-wrapper">
              <span class="page-title">webxcel-react-todo</span>
            </div>
          </div>
        </nav>

        <div class="container">
          <TodoList />
        </div>
      </div>
    )
  }
};


class ItemAddForm extends React.Component {
  constructor(props) {
    super(props);
    
    this.state = {
      value: ""
    };

    this.onAddButtonClicked = this.onAddButtonClicked.bind(this);
    this.handleInput = this.handleInput.bind(this);
  }

  onAddButtonClicked() {
    const value = this.state.value.trim();

    if (value.length > 0) {
      this.props.onItemAdded(value);
      this.setState({ value: "" });
    }
  }

  handleInput(event) {
    this.setState({ value: event.target.value });
  }

  render() {
    return (
      <div class="row">
        <div class="input-field col s8">
          <input placeholder="Type an item here and add it" value={this.state.value} onChange={this.handleInput} type="text" />
        </div>

        <div class="col s2 l3">
          <br />
          <a class="btn" onClick={this.onAddButtonClicked}>
            <i class="material-icons">add</i>
          </a>
        </div>
      </div>
    );
  }
}


class TodoList extends React.Component {
  constructor(props) {
    super(props);
    
    this.state = {
      items: []
    };

    this.addItem = this.addItem.bind(this);
    this.itemDeleted = this.itemDeleted.bind(this);
  }

  componentDidMount() {
    axios.get("/workbook/todo")
         .then(response => this.setState({
            items: response.data.map(item => {
              item.checked = item.checked == "TRUE";

              return item;
            })
         }));
  }

  addItem(title) {
    const id = uuid(),
          item = {
            id,
            title,
            checked: false
          };

    axios.post("/workbook/todo", item);

    this.state.items.push(item);
    this.setState(this.state);
  }

  itemCheckedChanged(id, checked) {
    axios.put(`/workbook/todo/${id}`, {
      id,
      checked
    });
  }

  itemDeleted(id) {
    const items = this.state.items.filter(item => item.id != id);

    this.setState({ items });
    axios.delete(`/workbook/todo/${id}`);
  }

  render() {
    const items = this.state.items.map(item => (
      <TodoItem id={item.id} title={item.title} checked={item.checked} onCheckedChanged={this.itemCheckedChanged} onDeleted={this.itemDeleted} />
    ));

    return (
      <div>
        <ItemAddForm onItemAdded={this.addItem} />

        <br />

        <div class="row">
          <div class="col s12">
            {items}
          </div>
        </div>
      </div>
    );
  }
}


class TodoItem extends React.Component {
  constructor(props) {
    super(props);
    
    this.state = {
      checked: props.checked
    };

    this.onChange = this.onChange.bind(this);
    this.onDeleted = this.onDeleted.bind(this);
  }

  onChange() {
    const id = this.props.id,
          checked = !this.state.checked;

    this.setState({ checked });
    this.props.onCheckedChanged(id, checked);
  }

  onDeleted() {
    const id = this.props.id;

    this.props.onDeleted(id);
  }

  render() {
    return (
      <div class="card">
        <div class="card-content row">
            <div class="col s10">
              <input type="checkbox" id={this.props.id} checked={this.state.checked} onChange={this.onChange} />
              <label for={this.props.id} class="black-text">
                <div class="title">{this.props.title}</div>
              </label>
            </div>
            <div class="col s2" onClick={this.onDeleted}>
              <i class="material-icons">delete</i>
            </div>
        </div>
      </div>
    );
  }
}


ReactDOM.render(
  <App />,
  document.getElementById("app")
);