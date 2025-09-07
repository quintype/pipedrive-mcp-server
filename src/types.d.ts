declare module 'pipedrive' {
  export namespace v1 {
    export interface ConfigurationParameters {
      apiKey?: string | Promise<string> | ((name: string) => string) | ((name: string) => Promise<string>);
    }

    export class Configuration {
      constructor(params: ConfigurationParameters);
    }

    export class DealsApi {
      constructor(configuration: Configuration);
      getDeals(params?: any): Promise<any>;
      getDeal(params: { id: number; [key: string]: any }): Promise<any>;
      searchDeals(params: { term: string; [key: string]: any }): Promise<any>;
    }

    export class PersonsApi {
      constructor(configuration: Configuration);
      getPersons(params?: any): Promise<any>;
      getPerson(params: { id: number; [key: string]: any }): Promise<any>;
      searchPersons(params: { term: string; [key: string]: any }): Promise<any>;
    }

    export class OrganizationsApi {
      constructor(configuration: Configuration);
      getOrganizations(params?: any): Promise<any>;
      getOrganization(params: { id: number; [key: string]: any }): Promise<any>;
      searchOrganizations(params: { term: string; [key: string]: any }): Promise<any>;
    }

    export class PipelinesApi {
      constructor(configuration: Configuration);
      getPipelines(params?: any): Promise<any>;
      getPipeline(params: { id: number; [key: string]: any }): Promise<any>;
      getPipelineStages(params: { id: number }): Promise<any>;
    }

    export class ItemSearchApi {
      constructor(configuration: Configuration);
      searchItem(params: { term: string, itemType?: string }): Promise<any>;
    }


    export class ActivitiesApi {
      constructor(configuration: Configuration);
      getActivities(params?: any): Promise<any>;
      addActivity(params: any): Promise<any>;
      getActivity(params: { id: number; [key: string]: any }): Promise<any>;
      updateActivity(params: { id: number; [key: string]: any }): Promise<any>;
      deleteActivity(params: { id: number }): Promise<any>;
    }

    export class ActivityFieldsApi {
      constructor(configuration: Configuration);
      getActivityFields(params?: any): Promise<any>;
      getActivityField(params: { id: number }): Promise<any>;
    }

    export class ActivityTypesApi {
      constructor(configuration: Configuration);
      getActivityTypes(params?: any): Promise<any>;
      getActivityType(params: { id: number }): Promise<any>;
    }

    export class CallLogsApi {
      constructor(configuration: Configuration);
      getCallLogs(params?: any): Promise<any>;
      addCallLog(params: any): Promise<any>;
      getCallLog(params: { id: string }): Promise<any>;
      updateCallLog(params: { id: string; [key: string]: any }): Promise<any>;
      deleteCallLog(params: { id: string }): Promise<any>;
    }

    export class CurrenciesApi {
      constructor(configuration: Configuration);
      getCurrencies(params?: any): Promise<any>;
    }

    export class DealFieldsApi {
      constructor(configuration: Configuration);
      getDealFields(params?: any): Promise<any>;
      getDealField(params: { id: number }): Promise<any>;
    }

    export class FilesApi {
      constructor(configuration: Configuration);
      getFiles(params?: any): Promise<any>;
      addFile(params: any): Promise<any>;
      getFile(params: { id: number }): Promise<any>;
      updateFile(params: { id: number; [key: string]: any }): Promise<any>;
      deleteFile(params: { id: number }): Promise<any>;
    }

    export class FiltersApi {
      constructor(configuration: Configuration);
      getFilters(params?: any): Promise<any>;
      addFilter(params: any): Promise<any>;
      getFilter(params: { id: number }): Promise<any>;
      updateFilter(params: { id: number; [key: string]: any }): Promise<any>;
      deleteFilter(params: { id: number }): Promise<any>;
    }

    export class GoalsApi {
      constructor(configuration: Configuration);
      getGoals(params?: any): Promise<any>;
      addGoal(params: any): Promise<any>;
      getGoal(params: { id: string }): Promise<any>;
      updateGoal(params: { id: string; [key: string]: any }): Promise<any>;
      deleteGoal(params: { id: string }): Promise<any>;
    }

    export class LeadLabelsApi {
      constructor(configuration: Configuration);
      getLeadLabels(params?: any): Promise<any>;
    }

    export class LeadSourcesApi {
      constructor(configuration: Configuration);
      getLeadSources(params?: any): Promise<any>;
    }
    export class LeadsApi {
      constructor(configuration: Configuration);
      getLeads(requestParameters?: {
        limit?: number;
        start?: number;
        owner_id?: number;
        person_id?: number;
        organization_id?: number;
        filter_id?: number;
        sort?: string;
      }): Promise<any>;
      getArchivedLeads(requestParameters?: {
        limit?: number;
        start?: number;
        owner_id?: number;
        person_id?: number;
        organization_id?: number;
        filter_id?: number;
        sort?: string;
      }): Promise<any>;
      getLead(requestParameters: { id: string }): Promise<any>;
      getLeadUsers(requestParameters: { id: string }): Promise<any>;
      addLead(requestParameters?: { AddLeadRequest?: any }): Promise<any>;
      updateLead(requestParameters: { id: string; UpdateLeadRequest?: any }): Promise<any>;
      deleteLead(requestParameters: { id: string }): Promise<any>;
      searchLeads(requestParameters: {
        term: string;
        fields?: "custom_fields" | "notes" | "title";
        exact_match?: boolean;
        person_id?: number;
        organization_id?: number;
        include_fields?: "lead.was_seen";
        start?: number;
        limit?: number;
      }): Promise<any>;
    }


    export class NoteFieldsApi {
      constructor(configuration: Configuration);
      getNoteFields(params?: any): Promise<any>;
      getNoteField(params: { id: number }): Promise<any>;
    }

    export class NotesApi {
      constructor(configuration: Configuration);
      getNotes(params?: any): Promise<any>;
      addNote(params: any): Promise<any>;
      getNote(params: { id: number }): Promise<any>;
      updateNote(params: { id: number; [key: string]: any }): Promise<any>;
      deleteNote(params: { id: number }): Promise<any>;
    }

    export class OrganizationFieldsApi {
      constructor(configuration: Configuration);
      getOrganizationFields(params?: any): Promise<any>;
      getOrganizationField(params: { id: number }): Promise<any>;
    }

    export class PersonFieldsApi {
      constructor(configuration: Configuration);
      getPersonFields(params?: any): Promise<any>;
      getPersonField(params: { id: number }): Promise<any>;
    }

    export class ProductFieldsApi {
      constructor(configuration: Configuration);
      getProductFields(params?: any): Promise<any>;
      getProductField(params: { id: number }): Promise<any>;
    }

    export class ProductsApi {
      constructor(configuration: Configuration);
      getProducts(params?: any): Promise<any>;
      addProduct(params: any): Promise<any>;
      getProduct(params: { id: number }): Promise<any>;
      updateProduct(params: { id: number; [key: string]: any }): Promise<any>;
      deleteProduct(params: { id: number }): Promise<any>;
    }

    export class StagesApi {
      constructor(configuration: Configuration);
      getStages(params?: any): Promise<any>;
      getStage(params: { id: number }): Promise<any>;
    }

    export class UsersApi {
      constructor(configuration: Configuration);
      getUsers(params?: any): Promise<any>;
      addUser(params: any): Promise<any>;
      getUser(params: { id: number }): Promise<any>;
      updateUser(params: { id: number; [key: string]: any }): Promise<any>;
      deleteUser(params: { id: number }): Promise<any>;
      getCurrentUser(params?: any): Promise<any>;
    }

    export class WebhooksApi {
      constructor(configuration: Configuration);
      getWebhooks(params?: any): Promise<any>;
      addWebhook(params: any): Promise<any>;
      getWebhook(params: { id: number }): Promise<any>;
      updateWebhook(params: { id: number; [key: string]: any }): Promise<any>;
      deleteWebhook(params: { id: number }): Promise<any>;
    }
  }

  export namespace v2 {
    export interface ConfigurationParameters {
      apiKey?: string | Promise<string> | ((name: string) => string) | ((name: string) => Promise<string>);
    }

    export class Configuration {
      constructor(params: ConfigurationParameters);
    }

    export class ActivitiesApi {
      constructor(configuration: Configuration);
      getActivities(params?: any): Promise<any>;
      addActivity(params: any): Promise<any>;
      getActivity(params: { id: number }): Promise<any>;
      updateActivity(params: { id: number; [key: string]: any }): Promise<any>;
      deleteActivity(params: { id: number }): Promise<any>;
    }

    export class BetaApi {
      constructor(configuration: Configuration);
      // Beta methods - experimental features
    }

    export class DealsApi {
      constructor(configuration: Configuration);
      getDeals(params?: any): Promise<any>;
      getDeal(params: { id: number }): Promise<any>;
      searchDeals(params: { term: string; [key: string]: any }): Promise<any>;
    }

    export class ItemSearchApi {
      constructor(configuration: Configuration);
      searchItem(params: { term: string, itemType?: string }): Promise<any>;
    }
    export class LeadsApi {
      constructor(configuration: Configuration);
      searchLeads(params: { term: string }): Promise<any>;
    }


    export class OrganizationsApi {
      constructor(configuration: Configuration);
      getOrganizations(params?: any): Promise<any>;
      getOrganization(params: { id: number }): Promise<any>;
      searchOrganizations(params: { term: string; [key: string]: any }): Promise<any>;
    }

    export class PersonsApi {
      constructor(configuration: Configuration);
      getPersons(params?: any): Promise<any>;
      getPerson(params: { id: number }): Promise<any>;
      searchPersons(params: { term: string; [key: string]: any }): Promise<any>;
    }

    export class PipelinesApi {
      constructor(configuration: Configuration);
      getPipelines(params?: any): Promise<any>;
      getPipeline(params: { id: number }): Promise<any>;
    }

    export class ProductsApi {
      constructor(configuration: Configuration);
      getProducts(params?: any): Promise<any>;
      addProduct(params: any): Promise<any>;
      getProduct(params: { id: number }): Promise<any>;
      updateProduct(params: { id: number; [key: string]: any }): Promise<any>;
      deleteProduct(params: { id: number }): Promise<any>;
      searchProducts(params: { term: string; [key: string]: any }): Promise<any>;
    }

    export class StagesApi {
      constructor(configuration: Configuration);
      getStages(params?: any): Promise<any>;
      getStage(params: { id: number }): Promise<any>;
    }

    export class UsersApi {
      constructor(configuration: Configuration);
      getUserFollowers(params: { id: number }): Promise<any>;
    }
  }
} 