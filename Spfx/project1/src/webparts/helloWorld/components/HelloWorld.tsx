import * as React from 'react';
import styles from './HelloWorld.module.scss';
import type { IHelloWorldProps } from './IHelloWorldProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface IEmployee {
  Id: number;
  Title: string;
  EmployeeID?: string;
  EmployeeName?: string;
  ExperienceInYears?: string;
  Grade?: string;
  Role?: string;
  Team?: string;
}

interface IEmployeeDraft {
  Title: string;
  EmployeeID: string;
  EmployeeName: string;
  ExperienceInYears: string;
  Grade: string;
  Role: string;
  Team: string;
}

interface IListField {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Choices?: string[] | { results: string[] };
}

interface ICollectionResponse<T> {
  value?: T[];
  d?: { results: T[] };
}

interface IHelloWorldState {
  employees: IEmployee[];
  listFields: IListField[];
  draft: IEmployeeDraft;
  editingId?: number;
  formOpen: boolean;
  search: string;
  loading: boolean;
  saving: boolean;
  error: string;
}

const emptyDraft: IEmployeeDraft = {
  Title: '',
  EmployeeID: '',
  EmployeeName: '',
  ExperienceInYears: '',
  Grade: '',
  Role: '',
  Team: ''
};

const fields: Array<{ name: keyof IEmployeeDraft; label: string }> = [
  { name: 'Title', label: 'Title' },
  { name: 'EmployeeID', label: 'Employee ID' },
  { name: 'EmployeeName', label: 'Employee name' },
  { name: 'ExperienceInYears', label: 'Experience in years' },
  { name: 'Grade', label: 'Grade' },
  { name: 'Role', label: 'Role' },
  { name: 'Team', label: 'Team' }
];

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {
  public state: IHelloWorldState = {
    employees: [],
    listFields: [],
    draft: { ...emptyDraft },
    formOpen: false,
    search: '',
    loading: true,
    saving: false,
    error: ''
  };

  public componentDidMount(): void {
    this.loadData().catch(() => undefined);
  }

  private get listApiUrl(): string {
    return `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'18f61653-649d-4e87-8bb6-21008c91ff61')`;
  }

  private async readJson<T>(response: SPHttpClientResponse): Promise<T> {
    if (!response.ok) {
      const detail = await response.text();
      throw new Error(detail || `${response.status} ${response.statusText}`);
    }
    return response.json();
  }

  private async loadData(): Promise<void> {
    this.setState({ loading: true, error: '' });
    try {
      const [itemsResponse, fieldsResponse] = await Promise.all([
        this.props.context.spHttpClient.get(
          `${this.listApiUrl}/items?$select=Id,Title,EmployeeID,EmployeeName,ExperienceInYears,Grade,Role,Team&$orderby=Modified desc`,
          SPHttpClient.configurations.v1
        ),
        this.props.context.spHttpClient.get(
          `${this.listApiUrl}/fields?$select=InternalName,Title,TypeAsString,Choices&$filter=Hidden eq false`,
          SPHttpClient.configurations.v1
        ),
      ]);
      const items = await this.readJson<ICollectionResponse<IEmployee>>(itemsResponse);
      const listFields = await this.readJson<ICollectionResponse<IListField>>(fieldsResponse);
      this.setState({
        employees: items.value || items.d?.results || [],
        listFields: listFields.value || listFields.d?.results || [],
        loading: false
      });
    } catch (error) {
      this.setState({ loading: false, error: this.getErrorMessage(error) });
    }
  }

  private getErrorMessage(error: unknown): string {
    return error instanceof Error ? error.message : 'Something went wrong while connecting to the list.';
  }

  private getChoiceOptions(fieldName: keyof IEmployeeDraft): string[] {
    const field = this.state.listFields.filter(item => item.InternalName === fieldName)[0];
    if (Array.isArray(field?.Choices)) {
      return field.Choices;
    }
    return field?.Choices?.results || [];
  }

  private beginCreate = (): void => {
    this.setState({ draft: { ...emptyDraft }, editingId: undefined, formOpen: true, error: '' });
  };

  private beginEdit = (employee: IEmployee): void => {
    this.setState({
      draft: {
        Title: employee.Title || '',
        EmployeeID: employee.EmployeeID || '',
        EmployeeName: employee.EmployeeName || '',
        ExperienceInYears: employee.ExperienceInYears || '',
        Grade: employee.Grade || '',
        Role: employee.Role || '',
        Team: employee.Team || ''
      },
      editingId: employee.Id,
      formOpen: true,
      error: ''
    });
  };

  private closeForm = (): void => {
    this.setState({ editingId: undefined, formOpen: false, draft: { ...emptyDraft }, saving: false });
  };

  private updateDraft = (field: keyof IEmployeeDraft, value: string): void => {
    this.setState(previous => ({ draft: { ...previous.draft, [field]: value } }));
  };

  private saveEmployee = async (event: React.FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();
    this.setState({ saving: true, error: '' });
    const isEditing = this.state.editingId !== undefined;
    const url = isEditing ? `${this.listApiUrl}/items(${this.state.editingId})` : `${this.listApiUrl}/items`;
    try {
      const response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': '',
          ...(isEditing ? { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' } : {})
        },
        body: JSON.stringify(this.state.draft)
      });
      if (!response.ok) {
        throw new Error(await response.text() || `${response.status} ${response.statusText}`);
      }
      this.closeForm();
      await this.loadData();
    } catch (error) {
      this.setState({ error: this.getErrorMessage(error), saving: false });
    }
  };

  private deleteEmployee = async (employee: IEmployee): Promise<void> => {
    if (!window.confirm(`Delete ${employee.EmployeeName || employee.Title}?`)) {
      return;
    }
    this.setState({ error: '' });
    try {
      const response = await this.props.context.spHttpClient.post(
        `${this.listApiUrl}/items(${employee.Id})`,
        SPHttpClient.configurations.v1,
        { headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' } }
      );
      if (!response.ok) {
        throw new Error(await response.text() || `${response.status} ${response.statusText}`);
      }
      await this.loadData();
    } catch (error) {
      this.setState({ error: this.getErrorMessage(error) });
    }
  };

  public render(): React.ReactElement<IHelloWorldProps> {
    const { employees, draft, editingId, formOpen, search, loading, saving, error } = this.state;
    const visibleEmployees = employees.filter(employee =>
      [employee.Title, employee.EmployeeID, employee.EmployeeName, employee.Role, employee.Team]
        .some(value => value?.toLowerCase().includes(search.toLowerCase()))
    );

    return (
      <section className={styles.employeeApp}>
        <header className={styles.header}>
          <div>
            <p className={styles.eyebrow}>PEOPLE DIRECTORY</p>
            <h2>Employee details</h2>
          </div>
          <button className={styles.primaryButton} type="button" onClick={this.beginCreate} disabled={loading}>+ Add employee</button>
        </header>

        <div className={styles.toolbar}>
          <label className={styles.search}>
            <span aria-hidden="true">⌕</span>
            <input
              type="search"
              value={search}
              onChange={event => this.setState({ search: event.currentTarget.value })}
              placeholder="Search employees"
              aria-label="Search employees"
            />
          </label>
          <span className={styles.resultCount}>{visibleEmployees.length} {visibleEmployees.length === 1 ? 'employee' : 'employees'}</span>
        </div>

        {error && <div className={styles.error} role="alert">{error}</div>}

        <div className={styles.tableWrap}>
          <table className={styles.employeeTable}>
            <thead>
              <tr>
                <th>Employee</th><th>Employee ID</th><th>Experience</th><th>Grade</th><th>Role</th><th>Team</th><th><span className={styles.visuallyHidden}>Actions</span></th>
              </tr>
            </thead>
            <tbody>
              {loading ? (
                <tr><td colSpan={7} className={styles.emptyState}>Loading employees...</td></tr>
              ) : visibleEmployees.length ? visibleEmployees.map(employee => (
                <tr key={employee.Id}>
                  <td><strong>{employee.EmployeeName || employee.Title}</strong><span className={styles.subText}>{employee.Title}</span></td>
                  <td>{employee.EmployeeID || '—'}</td>
                  <td>{employee.ExperienceInYears || '—'}</td>
                  <td>{employee.Grade || '—'}</td>
                  <td>{employee.Role || '—'}</td>
                  <td>{employee.Team || '—'}</td>
                  <td className={styles.actions}>
                    <button type="button" onClick={() => this.beginEdit(employee)} aria-label={`Edit ${employee.EmployeeName || employee.Title}`}>Edit</button>
                    <button type="button" onClick={() => this.deleteEmployee(employee).catch(() => undefined)} aria-label={`Delete ${employee.EmployeeName || employee.Title}`}>Delete</button>
                  </td>
                </tr>
              )) : (
                <tr><td colSpan={7} className={styles.emptyState}>{search ? 'No employees match your search.' : 'No employees yet. Add the first employee to get started.'}</td></tr>
              )}
            </tbody>
          </table>
        </div>

        {formOpen ? (
          <div className={styles.dialogBackdrop} role="presentation" onMouseDown={event => { if (event.target === event.currentTarget) this.closeForm(); }}>
            <section className={styles.dialog} role="dialog" aria-modal="true" aria-labelledby="employee-form-title">
              <div className={styles.dialogHeader}>
                <div><p className={styles.eyebrow}>{editingId !== undefined ? 'UPDATE RECORD' : 'NEW RECORD'}</p><h3 id="employee-form-title">{editingId !== undefined ? 'Edit employee' : 'Add employee'}</h3></div>
                <button className={styles.closeButton} type="button" onClick={this.closeForm} aria-label="Close">×</button>
              </div>
              <form onSubmit={event => this.saveEmployee(event).catch(() => undefined)}>
                <div className={styles.formGrid}>
                  {fields.map(field => {
                    const isChoice = ['ExperienceInYears', 'Grade', 'Role', 'Team'].indexOf(field.name) !== -1;
                    const options = this.getChoiceOptions(field.name);
                    return (
                      <label className={field.name === 'Title' || field.name === 'EmployeeName' ? styles.fullField : ''} key={field.name}>
                        <span>{field.label}{field.name === 'Title' && <b> *</b>}</span>
                        {isChoice ? (
                          <select value={draft[field.name]} onChange={event => this.updateDraft(field.name, event.currentTarget.value)}>
                            <option value="">Select {field.label.toLowerCase()}</option>
                            {options.map(option => <option key={option} value={option}>{option}</option>)}
                          </select>
                        ) : (
                          <input
                            required={field.name === 'Title'}
                            value={draft[field.name]}
                            onChange={event => this.updateDraft(field.name, event.currentTarget.value)}
                          />
                        )}
                      </label>
                    );
                  })}
                </div>
                <div className={styles.formActions}>
                  <button className={styles.secondaryButton} type="button" onClick={this.closeForm}>Cancel</button>
                  <button className={styles.primaryButton} type="submit" disabled={saving}>{saving ? 'Saving...' : editingId !== undefined ? 'Save changes' : 'Create employee'}</button>
                </div>
              </form>
            </section>
          </div>
        ) : null}
      </section>
    );
  }
}
