import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import * as dotenv from "dotenv";
import * as pipedrive from "pipedrive";
import { z } from "zod";

// Type for error handling
interface ErrorWithMessage {
  message: string;
}

function isErrorWithMessage(error: unknown): error is ErrorWithMessage {
  return (
    typeof error === "object" &&
    error !== null &&
    "message" in error &&
    typeof (error as Record<string, unknown>).message === "string"
  );
}

function getErrorMessage(error: unknown): string {
  if (isErrorWithMessage(error)) {
    return error.message;
  }
  return String(error);
}

// Load environment variables
dotenv.config();

// Check for required environment variables
if (!process.env.PIPEDRIVE_API_TOKEN) {
  console.error("ERROR: PIPEDRIVE_API_TOKEN environment variable is required");
  process.exit(1);
}

// Configure Pipedrive API authentication for both v1 and v2
const v1Configuration = new pipedrive.v1.Configuration({
  apiKey: process.env.PIPEDRIVE_API_TOKEN,
});

const v2Configuration = new pipedrive.v2.Configuration({
  apiKey: process.env.PIPEDRIVE_API_TOKEN,
});


// Helper function to remove custom_fields from response data
function cleanResponseData(data: any): any {
  if (!data) return data;
  
  // If it's an array, clean each item
  if (Array.isArray(data)) {
    return data.map(item => cleanResponseData(item));
  }
  
  // If it's an object, remove custom_fields key and clean nested objects
  if (typeof data === 'object') {
    const cleaned = { ...data };
    
    // Remove custom_fields key
    if ('custom_fields' in cleaned) {
      delete cleaned.custom_fields;
    }
    
    // Clean nested objects
    Object.keys(cleaned).forEach(key => {
      if (typeof cleaned[key] === 'object' && cleaned[key] !== null) {
        cleaned[key] = cleanResponseData(cleaned[key]);
      }
    });
    
    return cleaned;
  }
  
  return data;
}

// Initialize Pipedrive API clients - using v2 where available, v1 for others
// v2 APIs (modern, preferred where available)
const activitiesApi = new pipedrive.v2.ActivitiesApi(v2Configuration);
const dealsApi = new pipedrive.v2.DealsApi(v2Configuration);
const itemSearchApi = new pipedrive.v2.ItemSearchApi(v2Configuration);
// LeadsApi is only available in v1 - moved to v1 APIs section
const organizationsApi = new pipedrive.v2.OrganizationsApi(v2Configuration);
const personsApi = new pipedrive.v2.PersonsApi(v2Configuration);
const pipelinesApi = new pipedrive.v2.PipelinesApi(v2Configuration);
const productsApi = new pipedrive.v2.ProductsApi(v2Configuration);
const stagesApi = new pipedrive.v2.StagesApi(v2Configuration);
const usersApi = new pipedrive.v1.UsersApi(v1Configuration); // v2 only has getUserFollowers
const betaApi = new pipedrive.v2.BetaApi(v2Configuration);

// v1 APIs (comprehensive coverage for specialized endpoints)
const activityFieldsApi = new pipedrive.v1.ActivityFieldsApi(v1Configuration);
const activityTypesApi = new pipedrive.v1.ActivityTypesApi(v1Configuration);
const callLogsApi = new pipedrive.v1.CallLogsApi(v1Configuration);
const currenciesApi = new pipedrive.v1.CurrenciesApi(v1Configuration);
const dealFieldsApi = new pipedrive.v1.DealFieldsApi(v1Configuration);
const filesApi = new pipedrive.v1.FilesApi(v1Configuration);
const filtersApi = new pipedrive.v1.FiltersApi(v1Configuration);
const goalsApi = new pipedrive.v1.GoalsApi(v1Configuration);
const leadLabelsApi = new pipedrive.v1.LeadLabelsApi(v1Configuration);
const leadSourcesApi = new pipedrive.v1.LeadSourcesApi(v1Configuration);
const leadsV1Api = new pipedrive.v1.LeadsApi(v1Configuration);
const noteFieldsApi = new pipedrive.v1.NoteFieldsApi(v1Configuration);
const notesApi = new pipedrive.v1.NotesApi(v1Configuration);
const organizationFieldsApi = new pipedrive.v1.OrganizationFieldsApi(v1Configuration);
const personFieldsApi = new pipedrive.v1.PersonFieldsApi(v1Configuration);
const productFieldsApi = new pipedrive.v1.ProductFieldsApi(v1Configuration);
const webhooksApi = new pipedrive.v1.WebhooksApi(v1Configuration);

// Create MCP server
const server = new McpServer({
  name: "pipedrive-mcp-server",
  version: "1.0.0",
  capabilities: {
    resources: {},
    tools: {},
    prompts: {},
  },
});

// === TOOLS ===

// Get all deals
server.tool(
  "get-deals",
  "Retrieve all sales deals from Pipedrive CRM with advanced filtering by owner, person, organization, pipeline, stage, status, and date ranges. Includes deal amounts, stages, participants, and activities count.",
  {
    filter_id: z.number().optional().describe("Filter ID to match deals"),
    owner_id: z
      .number()
      .optional()
      .describe("User ID to filter deals by owner"),
    person_id: z.number().optional().describe("Person ID to filter deals"),
    org_id: z.number().optional().describe("Organization ID to filter deals"),
    pipeline_id: z.number().optional().describe("Pipeline ID to filter deals"),
    stage_id: z.number().optional().describe("Stage ID to filter deals"),
    status: z
      .enum(["open", "won", "lost", "deleted"])
      .optional()
      .describe("Deal status filter"),
    updated_since: z
      .string()
      .optional()
      .describe("RFC3339 datetime to filter deals updated after this time"),
    updated_until: z
      .string()
      .optional()
      .describe("RFC3339 datetime to filter deals updated before this time"),
    sort_by: z
      .enum(["id", "update_time", "add_time"])
      .optional()
      .describe("Field to sort by"),
    sort_direction: z
      .enum(["asc", "desc"])
      .optional()
      .describe("Sort direction"),
    include_fields: z
      .string()
      .optional()
      .describe("Comma-separated list of additional fields to include"),
    custom_fields: z
      .string()
      .optional()
      .describe("Comma-separated list of custom field keys"),
    limit: z
      .number()
      .max(500)
      .optional()
      .describe("Number of deals to return (max 500)"),
    cursor: z.string().optional().describe("Pagination cursor for next page"),
  },
  async (params) => {
    try {
      // Filter out undefined values and build query parameters
      const queryParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );


      const response = await dealsApi.getDeals(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching deals:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching deals: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get deal by ID
server.tool(
  "get-deal",
  "Fetch detailed information about a single sales deal using its unique Pipedrive ID. Returns deal amount, stage, owner, associated contacts, and all deal metadata.",
  {
    dealId: z.number().describe("Pipedrive deal ID"),
    include_fields: z
      .string()
      .optional()
      .describe("Comma-separated list of additional fields to include"),
    custom_fields: z
      .string()
      .optional()
      .describe("Comma-separated list of custom field keys"),
  },
  async ({ dealId, include_fields, custom_fields }) => {
    try {
      const queryParams: any = { id: dealId };
      
      if (include_fields) queryParams.include_fields = include_fields;
      
      if (custom_fields) queryParams.custom_fields = custom_fields;

      const response = await dealsApi.getDeal(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching deal ${dealId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching deal ${dealId}: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Search deals
server.tool(
  "search-deals",
  "Search for sales deals by keywords, company names, or deal titles. Find deals across all your pipelines with flexible text-based search and filtering options.",
  {
    term: z
      .string()
      .min(2)
      .describe("Search term for deals (minimum 2 characters)"),
    fields: z
      .string()
      .optional()
      .describe("Searchable field types to search in"),
    exact_match: z
      .boolean()
      .optional()
      .describe("Whether to use exact matching"),
    person_id: z.number().optional().describe("Filter by person ID"),
    organization_id: z
      .number()
      .optional()
      .describe("Filter by organization ID"),
    status: z
      .enum(["open", "won", "lost", "deleted"])
      .optional()
      .describe("Deal status filter"),
    include_fields: z
      .string()
      .optional()
      .describe("Comma-separated list of additional fields to include"),
    limit: z
      .number()
      .max(500)
      .optional()
      .describe("Number of results to return (max 500)"),
    cursor: z.string().optional().describe("Pagination cursor for next page"),
  },
  async (params) => {
    try {
      // Pass search parameters directly
      const searchParams = params;

      // Filter out undefined values but ensure required fields are preserved
      const queryParams = Object.fromEntries(
        Object.entries(searchParams).filter(
          ([_, value]) => value !== undefined
        )
      ) as typeof searchParams;


      const response = await dealsApi.searchDeals(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error searching deals with term "${params.term}":`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error searching deals: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get all persons
server.tool(
  "get-persons",
  "Retrieve all individual contacts and people from Pipedrive CRM. Includes contact details, phone numbers, emails, associated deals, and relationship information. These are individual people, not company users or owners.",
  {
    filter_id: z.number().optional().describe("Filter ID to match persons"),
    owner_id: z
      .number()
      .optional()
      .describe("User ID to filter persons by owner"),
    org_id: z.number().optional().describe("Organization ID to filter persons"),
    updated_since: z
      .string()
      .optional()
      .describe("RFC3339 datetime to filter persons updated after this time"),
    updated_until: z
      .string()
      .optional()
      .describe("RFC3339 datetime to filter persons updated before this time"),
    sort_by: z
      .enum(["id", "update_time", "add_time"])
      .optional()
      .describe("Field to sort by"),
    sort_direction: z
      .enum(["asc", "desc"])
      .optional()
      .describe("Sort direction"),
    include_fields: z
      .string()
      .optional()
      .describe(
        "Comma-separated list of additional fields like marketing_status"
      ),
    custom_fields: z
      .string()
      .optional()
      .describe("Comma-separated list of custom field keys (max 15)"),
    limit: z
      .number()
      .max(500)
      .optional()
      .describe("Number of persons to return (max 500, default 100)"),
    cursor: z.string().optional().describe("Pagination cursor for next page"),
  },
  async (params) => {
    try {
      // Filter out undefined values
      const queryParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );


      const response = await personsApi.getPersons(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching persons:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching persons: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get person by ID
server.tool(
  "get-person",
  "Fetch detailed contact information for a single person/individual contact using their Pipedrive person ID. Returns personal details, communication history, and associated deals.",
  {
    personId: z.number().describe("Pipedrive person ID"),
    include_fields: z
      .string()
      .optional()
      .describe("Comma-separated list of additional fields to include"),
    custom_fields: z
      .string()
      .optional()
      .describe("Comma-separated list of custom field keys"),
  },
  async ({ personId, include_fields, custom_fields }) => {
    try {
      const queryParams: any = { id: personId };
      
      if (include_fields) queryParams.include_fields = include_fields;
      
      if (custom_fields) queryParams.custom_fields = custom_fields;

      const response = await personsApi.getPerson(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching person ${personId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching person ${personId}: ${getErrorMessage(
              error
            )}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Search persons
server.tool(
  "search-persons",
  "Search for individual contacts and people by name, email, phone number, or other personal details. Find specific contacts in your Pipedrive database.",
  {
    term: z
      .string()
      .min(2)
      .describe("Search term for persons (minimum 2 characters)"),
    fields: z
      .string()
      .optional()
      .describe("Searchable field types to search in"),
    exact_match: z
      .boolean()
      .optional()
      .describe("Whether to use exact matching"),
    organization_id: z
      .number()
      .optional()
      .describe("Filter by organization ID"),
    limit: z
      .number()
      .max(500)
      .optional()
      .describe("Number of results to return (max 500)"),
    cursor: z.string().optional().describe("Pagination cursor for next page"),
  },
  async (params) => {
    try {
      // Pass search parameters directly
      const searchParams = params;

      // Filter out undefined values but ensure required fields are preserved
      const queryParams = Object.fromEntries(
        Object.entries(searchParams).filter(
          ([_, value]) => value !== undefined
        )
      ) as typeof searchParams;


      const response = await personsApi.searchPersons(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(
        `Error searching persons with term "${params.term}":`,
        error
      );
      return {
        content: [
          {
            type: "text",
            text: `Error searching persons: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get all organizations
server.tool(
  "get-organizations",
  "Retrieve all companies and organizations from Pipedrive CRM. Includes company details, addresses, associated contacts, deals, and organizational hierarchy information.",
  {
    filter_id: z
      .number()
      .optional()
      .describe("Filter ID to match organizations"),
    ids: z.string().optional().describe("Comma-separated entity IDs to fetch"),
    owner_id: z
      .number()
      .optional()
      .describe("User ID to filter organizations by owner"),
    updated_since: z
      .string()
      .optional()
      .describe("Organizations with update_time later than this time"),
    updated_until: z
      .string()
      .optional()
      .describe("Organizations with update_time earlier than this time"),
    sort_by: z
      .enum(["id", "update_time", "add_time"])
      .optional()
      .describe("Field to sort by"),
    sort_direction: z
      .enum(["asc", "desc"])
      .optional()
      .describe("Sort direction"),
    include_fields: z
      .string()
      .optional()
      .describe("Additional fields to include"),
    custom_fields: z
      .string()
      .optional()
      .describe("Custom field keys to include"),
    limit: z
      .number()
      .max(500)
      .optional()
      .describe("Number of entries to return (max 500)"),
    cursor: z.string().optional().describe("Pagination cursor"),
  },
  async (params) => {
    try {
      // Filter out undefined values
      const queryParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );


      const response = await organizationsApi.getOrganizations(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching organizations:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching organizations: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get organization by ID
server.tool(
  "get-organization",
  "Fetch detailed company information for a single organization using its Pipedrive organization ID. Returns company details, contacts, deals, and business relationships.",
  {
    organizationId: z.number().describe("Pipedrive organization ID"),
    include_fields: z
      .string()
      .optional()
      .describe("Additional fields to include"),
    custom_fields: z
      .string()
      .optional()
      .describe("Custom field keys to include"),
  },
  async ({ organizationId, include_fields, custom_fields }) => {
    try {
      const queryParams: any = { id: organizationId };
      
      if (include_fields) queryParams.include_fields = include_fields;
      
      if (custom_fields) queryParams.custom_fields = custom_fields;

      const response = await organizationsApi.getOrganization(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching organization ${organizationId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching organization ${organizationId}: ${getErrorMessage(
              error
            )}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Search organizations
server.tool(
  "search-organizations",
  "Search for companies and organizations by company name, domain, industry, or other business details. Find specific organizations in your CRM database.",
  {
    term: z
      .string()
      .min(2)
      .describe("Search term for organizations (minimum 2 characters)"),
    fields: z.string().optional().describe("Fields to search in"),
    exact_match: z
      .boolean()
      .optional()
      .describe("Whether to use exact matching"),
    limit: z
      .number()
      .max(500)
      .optional()
      .describe("Number of entries to return (max 500)"),
    cursor: z.string().optional().describe("Pagination cursor"),
  },
  async (params) => {
    try {
      // Pass search parameters directly
      const searchParams = params;

      // Filter out undefined values but ensure required fields are preserved
      const queryParams = Object.fromEntries(
        Object.entries(searchParams).filter(
          ([_, value]) => value !== undefined
        )
      ) as typeof searchParams;


      const response = await organizationsApi.searchOrganizations(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(
        `Error searching organizations with term "${params.term}":`,
        error
      );
      return {
        content: [
          {
            type: "text",
            text: `Error searching organizations: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get all pipelines
server.tool(
  "get-pipelines",
  "Retrieve all sales pipelines configured in your Pipedrive CRM. Shows pipeline names, stages, deal flow configuration, and pipeline-specific settings.",
  {
    sort_by: z
      .enum(["id", "update_time", "add_time"])
      .optional()
      .describe("Field to sort by (default: id)"),
    sort_direction: z
      .enum(["asc", "desc"])
      .optional()
      .describe("Sort direction (default: asc)"),
    limit: z
      .number()
      .max(500)
      .optional()
      .describe("Number of pipelines to return (max 500, default 100)"),
    cursor: z.string().optional().describe("Pagination cursor for next page"),
  },
  async (params) => {
    try {
      // Filter out undefined values
      const queryParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );

      const response = await pipelinesApi.getPipelines(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching pipelines:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching pipelines: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get pipeline by ID
server.tool(
  "get-pipeline",
  "Fetch detailed information about a single sales pipeline using its ID. Returns pipeline stages, configuration, and associated deals count.",
  {
    pipelineId: z.number().describe("Pipedrive pipeline ID"),
  },
  async ({ pipelineId }) => {
    try {
      const response = await pipelinesApi.getPipeline({ id: pipelineId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching pipeline ${pipelineId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching pipeline ${pipelineId}: ${getErrorMessage(
              error
            )}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get all stages
server.tool("get-stages", "Retrieve all pipeline stages across all sales pipelines. Shows stage names, win probability, order, and stage-specific configuration settings.", {}, async () => {
  try {
    // Since the stages are related to pipelines, we'll get all pipelines first
    const pipelinesResponse = await pipelinesApi.getPipelines();
    const pipelines = pipelinesResponse.data || [];

    // For each pipeline, fetch its stages
    const allStages = [];
    for (const pipeline of pipelines) {
      try {
        // This is using the API to get stages by pipeline ID
        const stagesResponse = await fetch(
          `https://api.pipedrive.com/api/v2/stages?pipeline_id=${pipeline.id}&api_token=${process.env.PIPEDRIVE_API_TOKEN}`
        );
        const stagesData = await stagesResponse.json();

        if (stagesData.success && stagesData.data) {
          const pipelineStages = stagesData.data.map((stage: any) => ({
            ...stage,
            pipeline_name: pipeline.name,
          }));
          allStages.push(...pipelineStages);
        }
      } catch (e) {
        console.error(`Error fetching stages for pipeline ${pipeline.id}:`, e);
      }
    }

    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(allStages, null, 2),
        },
      ],
    };
  } catch (error) {
    console.error("Error fetching stages:", error);
    return {
      content: [
        {
          type: "text",
          text: `Error fetching stages: ${getErrorMessage(error)}`,
        },
      ],
      isError: true,
    };
  }
});

// Search leads
server.tool(
  "search-leads",
  "Search for sales leads and prospects by keywords, company name, or contact details. Find potential customers in your lead database before they become deals.",
  {
    term: z.string().describe("Search term for leads"),
  },
  async ({ term }) => {
    try {
      // URL encode the search term
      const searchTerm = term;
      const response = await leadsV1Api.searchLeads({ term: searchTerm });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error searching leads with term "${term}":`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error searching leads: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Generic search across item types
server.tool(
  "search-all",
  "Perform universal search across all Pipedrive data types including deals, contacts (persons), companies (organizations), products, and notes. One search to find any CRM record.",
  {
    term: z.string().describe("Search term"),
    itemTypes: z
      .string()
      .optional()
      .describe(
        "Comma-separated list of item types to search (deal,person,organization,product,file,activity,lead)"
      ),
  },
  async ({ term, itemTypes }) => {
    try {
      // URL encode the search term
      const searchTerm = term;
      const itemType = itemTypes; // Just rename the parameter
      const response = await itemSearchApi.searchItem({
        term: searchTerm,
        itemType,
      });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error performing search with term "${term}":`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error performing search: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === ACTIVITIES API ===

// Get all activities
server.tool(
  "get-activities",
  "Retrieve all scheduled and completed activities (calls, meetings, emails, tasks) from Pipedrive. Filter by type, date, owner, or associated deals and contacts.",
  {
    filter_id: z.number().optional().describe("Filter ID to match activities"),
    user_id: z
      .number()
      .optional()
      .describe("User ID to filter activities by assignee"),
    deal_id: z.number().optional().describe("Deal ID to filter activities"),
    person_id: z.number().optional().describe("Person ID to filter activities"),
    org_id: z
      .number()
      .optional()
      .describe("Organization ID to filter activities"),
    type: z.string().optional().describe("Activity type"),
    updated_since: z
      .string()
      .optional()
      .describe(
        "RFC3339 datetime to filter activities updated after this time"
      ),
    updated_until: z
      .string()
      .optional()
      .describe(
        "RFC3339 datetime to filter activities updated before this time"
      ),
    start_date: z
      .string()
      .optional()
      .describe("Start date filter (YYYY-MM-DD)"),
    end_date: z.string().optional().describe("End date filter (YYYY-MM-DD)"),
    done: z
      .enum(["0", "1"])
      .optional()
      .describe("Filter by done status (0=not done, 1=done)"),
    limit: z
      .number()
      .max(500)
      .optional()
      .describe("Number of activities to return (max 500)"),
    cursor: z.string().optional().describe("Pagination cursor for next page"),
  },
  async (params) => {
    try {
      const queryParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );

      const response = await activitiesApi.getActivities(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching activities:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching activities: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Create activity
server.tool(
  "create-activity",
  "Schedule a new activity such as a call, meeting, email, or task in Pipedrive. Link activities to deals, contacts, or organizations for better CRM tracking.",
  {
    subject: z.string().describe("Activity subject"),
    type: z.string().optional().describe("Activity type"),
    deal_id: z.number().optional().describe("Deal ID to associate with"),
    person_id: z.number().optional().describe("Person ID to associate with"),
    org_id: z.number().optional().describe("Organization ID to associate with"),
    user_id: z.number().optional().describe("User ID to assign activity to"),
    due_date: z.string().optional().describe("Due date (YYYY-MM-DD)"),
    due_time: z.string().optional().describe("Due time (HH:MM)"),
    duration: z.string().optional().describe("Duration in format H:MM"),
    note: z.string().optional().describe("Activity note"),
    done: z
      .enum(["0", "1"])
      .optional()
      .describe("Mark as done (0=not done, 1=done)"),
  },
  async (params) => {
    try {
      const response = await activitiesApi.addActivity(params);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error creating activity:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error creating activity: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get activity by ID
server.tool(
  "get-activity",
  "Fetch detailed information about a single activity using its ID. Returns activity type, participants, notes, completion status, and linked CRM records.",
  {
    activityId: z.number().describe("Activity ID"),
  },
  async ({ activityId }) => {
    try {
      const response = await activitiesApi.getActivity({ id: activityId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching activity ${activityId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching activity ${activityId}: ${getErrorMessage(
              error
            )}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Update activity
server.tool(
  "update-activity",
  "Modify details of an existing activity such as changing the date, participants, notes, or completion status. Update calls, meetings, emails, or tasks.",
  {
    activityId: z.number().describe("Activity ID"),
    subject: z.string().optional().describe("Activity subject"),
    type: z.string().optional().describe("Activity type"),
    deal_id: z.number().optional().describe("Deal ID to associate with"),
    person_id: z.number().optional().describe("Person ID to associate with"),
    org_id: z.number().optional().describe("Organization ID to associate with"),
    user_id: z.number().optional().describe("User ID to assign activity to"),
    due_date: z.string().optional().describe("Due date (YYYY-MM-DD)"),
    due_time: z.string().optional().describe("Due time (HH:MM)"),
    duration: z.string().optional().describe("Duration in format H:MM"),
    note: z.string().optional().describe("Activity note"),
    done: z
      .enum(["0", "1"])
      .optional()
      .describe("Mark as done (0=not done, 1=done)"),
  },
  async ({ activityId, ...params }) => {
    try {
      const updateParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );

      const response = await activitiesApi.updateActivity({
        id: activityId,
        ...updateParams,
      });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error updating activity ${activityId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error updating activity ${activityId}: ${getErrorMessage(
              error
            )}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Delete activity
server.tool(
  "delete-activity",
  "Remove an activity from Pipedrive permanently. Delete scheduled or completed calls, meetings, emails, or tasks that are no longer needed.",
  {
    activityId: z.number().describe("Activity ID"),
  },
  async ({ activityId }) => {
    try {
      const response = await activitiesApi.deleteActivity({ id: activityId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error deleting activity ${activityId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error deleting activity ${activityId}: ${getErrorMessage(
              error
            )}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === ACTIVITY FIELDS API ===

// Get all activity fields
server.tool(
  "get-activity-fields",
  "Retrieve all available activity field definitions and metadata. Shows custom fields, field types, and configuration for activity forms and data structure.",
  {},
  async () => {
    try {
      const response = await activityFieldsApi.getActivityFields();
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching activity fields:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching activity fields: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get activity field by ID
server.tool(
  "get-activity-field",
  "Fetch detailed information about a single activity field configuration using its field ID. Returns field type, options, validation rules, and usage.",
  {
    fieldId: z.number().describe("Activity field ID"),
  },
  async ({ fieldId }) => {
    try {
      const response = await activityFieldsApi.getActivityField({
        id: fieldId,
      });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching activity field ${fieldId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching activity field ${fieldId}: ${getErrorMessage(
              error
            )}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === ACTIVITY TYPES API ===

// Get all activity types
server.tool(
  "get-activity-types",
  "Retrieve all configured activity types (call, meeting, email, lunch, etc.). Shows available activity categories and their icons, colors, and settings.",
  {},
  async () => {
    try {
      const response = await activityTypesApi.getActivityTypes();
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching activity types:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching activity types: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get activity type by ID
server.tool(
  "get-activity-type",
  "Fetch detailed information about a single activity type using its ID. Returns activity type configuration, default settings, and usage statistics.",
  {
    typeId: z.number().describe("Activity type ID"),
  },
  async ({ typeId }) => {
    try {
      const response = await activityTypesApi.getActivityType({ id: typeId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching activity type ${typeId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching activity type ${typeId}: ${getErrorMessage(
              error
            )}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === CURRENCIES API ===

// Get all currencies
server.tool(
  "get-currencies",
  "Retrieve all supported currencies configured in your Pipedrive account. Shows currency codes, symbols, exchange rates, and formatting preferences.",
  {},
  async () => {
    try {
      const response = await currenciesApi.getCurrencies();
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching currencies:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching currencies: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === USERS API ===

// Get all users
server.tool("get-users", "Retrieve all Pipedrive users, owners, and team members. These are your company's internal users who can own deals, manage pipelines, and access the CRM system. Use this to find deal owners, sales reps, and team members.", {}, async () => {
  try {
    const response = await usersApi.getUsers();
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(cleanResponseData(response.data), null, 2),
        },
      ],
    };
  } catch (error) {
    console.error("Error fetching users:", error);
    return {
      content: [
        {
          type: "text",
          text: `Error fetching users: ${getErrorMessage(error)}`,
        },
      ],
      isError: true,
    };
  }
});

// Get current user
server.tool(
  "get-current-user",
  "Fetch information about the currently authenticated Pipedrive user/owner. Returns user profile, permissions, and account settings for the API token owner.",
  {},
  async () => {
    try {
      const response = await usersApi.getCurrentUser();
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching current user:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching current user: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get user by ID
server.tool(
  "get-user",
  "Fetch detailed information about a specific Pipedrive user, owner, or team member using their user ID. Returns user profile, role, and access permissions.",
  {
    userId: z.number().describe("User ID"),
  },
  async ({ userId }) => {
    try {
      const response = await usersApi.getUser({ id: userId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching user ${userId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching user ${userId}: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === NOTES API ===

// Get all notes
server.tool(
  "get-notes",
  "Retrieve all notes and comments from Pipedrive CRM. Includes notes attached to deals, contacts, organizations, and standalone notes with filtering options.",
  {
    deal_id: z.number().optional().describe("Deal ID to filter notes"),
    person_id: z.number().optional().describe("Person ID to filter notes"),
    org_id: z.number().optional().describe("Organization ID to filter notes"),
    start: z.number().optional().describe("Pagination start"),
    limit: z.number().optional().describe("Items per page"),
  },
  async (params) => {
    try {
      const queryParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );

      const response = await notesApi.getNotes(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching notes:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching notes: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Create note
server.tool(
  "create-note",
  "Add a new note or comment in Pipedrive. Attach notes to deals, contacts, organizations, or create standalone notes for future reference and team collaboration.",
  {
    content: z.string().describe("Note content"),
    deal_id: z.number().optional().describe("Deal ID to associate with"),
    person_id: z.number().optional().describe("Person ID to associate with"),
    org_id: z.number().optional().describe("Organization ID to associate with"),
    pinned_to_deal_flag: z
      .enum(["0", "1"])
      .optional()
      .describe("Pin to deal (0=no, 1=yes)"),
    pinned_to_person_flag: z
      .enum(["0", "1"])
      .optional()
      .describe("Pin to person (0=no, 1=yes)"),
    pinned_to_organization_flag: z
      .enum(["0", "1"])
      .optional()
      .describe("Pin to organization (0=no, 1=yes)"),
  },
  async (params) => {
    try {
      const response = await notesApi.addNote(params);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error creating note:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error creating note: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get note by ID
server.tool(
  "get-note",
  "Fetch detailed information about a single note using its ID. Returns note content, attachments, creation date, author, and linked CRM records.",
  {
    noteId: z.number().describe("Note ID"),
  },
  async ({ noteId }) => {
    try {
      const response = await notesApi.getNote({ id: noteId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching note ${noteId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching note ${noteId}: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === PRODUCTS API ===

// Get all products
server.tool(
  "get-products",
  "Retrieve all products and services from your Pipedrive catalog. Includes product names, prices, descriptions, categories, and inventory information.",
  {
    filter_id: z.number().optional().describe("Filter ID to match products"),
    ids: z.string().optional().describe("Comma-separated product IDs to fetch"),
    first_char: z
      .string()
      .optional()
      .describe("Filter by first character of product name"),
    get_summary: z.boolean().optional().describe("Include summary in response"),
    start: z.number().optional().describe("Pagination start"),
    limit: z.number().optional().describe("Items per page"),
  },
  async (params) => {
    try {
      const queryParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );

      const response = await productsApi.getProducts(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching products:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching products: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Create product
server.tool(
  "create-product",
  "Add a new product or service to your Pipedrive catalog. Define product names, prices, descriptions, units, and categories for deal line items.",
  {
    name: z.string().describe("Product name"),
    code: z.string().optional().describe("Product code"),
    description: z.string().optional().describe("Product description"),
    unit: z.string().optional().describe("Product unit"),
    tax: z.number().optional().describe("Tax percentage"),
    category: z.string().optional().describe("Product category"),
    active_flag: z.boolean().optional().describe("Whether product is active"),
    visible_to: z
      .enum(["1", "3", "5", "7"])
      .optional()
      .describe("Visibility level"),
    owner_id: z.number().optional().describe("Product owner ID"),
    prices: z
      .array(
        z.object({
          price: z.number().describe("Price amount"),
          currency: z.string().describe("Currency code"),
          cost: z.number().optional().describe("Cost amount"),
          overhead_cost: z.number().optional().describe("Overhead cost"),
        })
      )
      .optional()
      .describe("Product prices"),
  },
  async (params) => {
    try {
      const response = await productsApi.addProduct(params);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error creating product:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error creating product: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get product by ID
server.tool(
  "get-product",
  "Fetch detailed information about a single product using its ID. Returns pricing, description, category, availability, and usage in deals.",
  {
    productId: z.number().describe("Product ID"),
  },
  async ({ productId }) => {
    try {
      const response = await productsApi.getProduct({ id: productId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching product ${productId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching product ${productId}: ${getErrorMessage(
              error
            )}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === STAGES API ===

// Get all stages
server.tool(
  "get-stages-v2",
  "Retrieve all pipeline stages using the advanced StagesApi. Shows stage details, win probabilities, rotting time, and stage-specific configuration settings.",
  {
    pipeline_id: z.number().optional().describe("Filter stages by pipeline ID"),
  },
  async (params) => {
    try {
      const queryParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );

      const response = await stagesApi.getStages(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching stages:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching stages: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get stage by ID
server.tool(
  "get-stage",
  "Fetch detailed information about a single pipeline stage using its ID. Returns stage configuration, deals count, win rate, and stage progression metrics.",
  {
    stageId: z.number().describe("Stage ID"),
  },
  async ({ stageId }) => {
    try {
      const response = await stagesApi.getStage({ id: stageId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching stage ${stageId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching stage ${stageId}: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === FIELDS APIs ===

// Get deal fields
server.tool(
  "get-deal-fields",
  "Retrieve all deal field definitions and custom fields. Shows field types, validation rules, options, and configuration for deal forms and data structure.",
  {},
  async () => {
    try {
      const response = await dealFieldsApi.getDealFields();
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching deal fields:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching deal fields: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get person fields
server.tool(
  "get-person-fields",
  "Retrieve all contact/person field definitions and custom fields. Shows field types, validation rules, and configuration for contact forms and data structure.",
  {},
  async () => {
    try {
      const response = await personFieldsApi.getPersonFields();
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching person fields:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching person fields: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get organization fields
server.tool(
  "get-organization-fields",
  "Retrieve all company/organization field definitions and custom fields. Shows field types, validation rules, and configuration for organization forms and data.",
  {},
  async () => {
    try {
      const response = await organizationFieldsApi.getOrganizationFields();
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching organization fields:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching organization fields: ${getErrorMessage(
              error
            )}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === CALL LOGS API ===

// Get all call logs
server.tool(
  "get-call-logs",
  "Retrieve all phone call records and logs from Pipedrive. Includes call duration, outcomes, participants, recordings, and notes from telephony integrations.",
  {
    start: z.number().optional().describe("Pagination start"),
    limit: z.number().optional().describe("Items per page"),
  },
  async (params) => {
    try {
      const queryParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );

      const response = await callLogsApi.getCallLogs(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching call logs:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching call logs: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Create call log
server.tool(
  "create-call-log",
  "Record a new phone call log in Pipedrive. Track call duration, outcome, participants, notes, and link calls to deals, contacts, or organizations.",
  {
    subject: z.string().describe("Call subject"),
    outcome: z.string().describe("Call outcome"),
    from_phone_number: z.string().optional().describe("From phone number"),
    to_phone_number: z.string().optional().describe("To phone number"),
    start_time: z.string().optional().describe("Call start time"),
    end_time: z.string().optional().describe("Call end time"),
    duration: z.number().optional().describe("Call duration in seconds"),
    activity_id: z.number().optional().describe("Related activity ID"),
    person_id: z.number().optional().describe("Person ID"),
    org_id: z.number().optional().describe("Organization ID"),
    deal_id: z.number().optional().describe("Deal ID"),
  },
  async (params) => {
    try {
      const response = await callLogsApi.addCallLog(params);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error creating call log:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error creating call log: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get call log by ID
server.tool(
  "get-call-log",
  "Fetch detailed information about a single phone call record using its ID. Returns call details, duration, outcome, participants, and linked records.",
  {
    callLogId: z.string().describe("Call log ID"),
  },
  async ({ callLogId }) => {
    try {
      const response = await callLogsApi.getCallLog({ id: callLogId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching call log ${callLogId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching call log ${callLogId}: ${getErrorMessage(
              error
            )}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === FILES API ===

// Get all files
server.tool(
  "get-files",
  "Retrieve all uploaded files and documents from Pipedrive. Includes files attached to deals, contacts, organizations, and standalone document storage.",
  {
    start: z.number().optional().describe("Pagination start"),
    limit: z.number().optional().describe("Items per page"),
    include_deleted_files: z
      .boolean()
      .optional()
      .describe("Include deleted files"),
    sort: z.string().optional().describe("Sorting field"),
  },
  async (params) => {
    try {
      const queryParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );

      const response = await filesApi.getFiles(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching files:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching files: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get file by ID
server.tool(
  "get-file",
  "Fetch detailed information about a single uploaded file using its ID. Returns file metadata, size, type, upload date, and associated CRM records.",
  {
    fileId: z.number().describe("File ID"),
  },
  async ({ fileId }) => {
    try {
      const response = await filesApi.getFile({ id: fileId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching file ${fileId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching file ${fileId}: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === FILTERS API ===

// Get all filters
server.tool(
  "get-filters",
  "Retrieve all saved filters and views configured in Pipedrive. Shows custom filters for deals, contacts, organizations, and their filter criteria.",
  {
    type: z
      .enum(["deals", "people", "organizations", "products", "activities"])
      .optional()
      .describe("Filter type"),
  },
  async (params) => {
    try {
      const queryParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );

      const response = await filtersApi.getFilters(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching filters:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching filters: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get filter by ID
server.tool(
  "get-filter",
  "Fetch detailed information about a single saved filter using its ID. Returns filter criteria, conditions, and usage statistics.",
  {
    filterId: z.number().describe("Filter ID"),
  },
  async ({ filterId }) => {
    try {
      const response = await filtersApi.getFilter({ id: filterId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching filter ${filterId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching filter ${filterId}: ${getErrorMessage(
              error
            )}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === GOALS API ===

// Get all goals
server.tool(
  "get-goals",
  "Retrieve all sales goals and targets configured in Pipedrive. Shows revenue targets, activity goals, performance metrics, and goal progress tracking.",
  {
    everyone: z.boolean().optional().describe("Get goals for everyone"),
    user_id: z.number().optional().describe("User ID to get goals for"),
  },
  async (params) => {
    try {
      const queryParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );

      const response = await goalsApi.getGoals(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching goals:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching goals: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get goal by ID
server.tool(
  "get-goal",
  "Fetch detailed information about a single sales goal using its ID. Returns target values, current progress, timeframes, and assigned users.",
  {
    goalId: z.string().describe("Goal ID"),
  },
  async ({ goalId }) => {
    try {
      const response = await goalsApi.getGoal({ id: goalId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching goal ${goalId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching goal ${goalId}: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === WEBHOOKS API ===

// Get all webhooks
server.tool("get-webhooks", "Retrieve all configured webhooks and API integrations from Pipedrive. Shows webhook URLs, events, triggers, and integration status.", {}, async () => {
  try {
    const response = await webhooksApi.getWebhooks();
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify(cleanResponseData(response.data), null, 2),
        },
      ],
    };
  } catch (error) {
    console.error("Error fetching webhooks:", error);
    return {
      content: [
        {
          type: "text",
          text: `Error fetching webhooks: ${getErrorMessage(error)}`,
        },
      ],
      isError: true,
    };
  }
});

// Create webhook
server.tool(
  "create-webhook",
  "Configure a new webhook integration in Pipedrive. Set up real-time notifications for CRM events like new deals, updated contacts, or completed activities.",
  {
    subscription_url: z.string().describe("Webhook subscription URL"),
    event_action: z
      .string()
      .describe("Event action (e.g., added, updated, deleted)"),
    event_object: z
      .string()
      .describe("Event object (e.g., deal, person, organization)"),
    user_id: z
      .number()
      .optional()
      .describe("User ID to associate webhook with"),
    http_auth_user: z.string().optional().describe("HTTP auth username"),
    http_auth_password: z.string().optional().describe("HTTP auth password"),
  },
  async (params) => {
    try {
      const response = await webhooksApi.addWebhook(params);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error creating webhook:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error creating webhook: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get webhook by ID
server.tool(
  "get-webhook",
  "Fetch detailed information about a single webhook configuration using its ID. Returns webhook URL, events, triggers, and delivery status.",
  {
    webhookId: z.number().describe("Webhook ID"),
  },
  async ({ webhookId }) => {
    try {
      const response = await webhooksApi.getWebhook({ id: webhookId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching webhook ${webhookId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching webhook ${webhookId}: ${getErrorMessage(
              error
            )}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === COMPREHENSIVE LEADS API ===

// Get all leads
server.tool(
  "get-leads",
  "Retrieve all sales leads and prospects from Pipedrive CRM. Leads are potential customers before they become deals. Filter by owner, person, organization, or custom filters.",
  {
    limit: z.number().optional().describe("For pagination, the limit of entries to be returned. If not provided, 100 items will be returned."),
    start: z.number().optional().describe("For pagination, the position that represents the first result for the page"),
    owner_id: z.number().optional().describe("If supplied, only leads matching the given user will be returned. However, filter_id takes precedence over owner_id when supplied."),
    person_id: z.number().optional().describe("If supplied, only leads matching the given person will be returned. However, filter_id takes precedence over person_id when supplied."),
    organization_id: z.number().optional().describe("If supplied, only leads matching the given organization will be returned. However, filter_id takes precedence over organization_id when supplied."),
    filter_id: z.number().optional().describe("The ID of the filter to use"),
    sort: z.string().optional().describe("The field names and sorting mode separated by a comma (field_name_1 ASC, field_name_2 DESC). Supported fields: id, title, owner_id, creator_id, was_seen, expected_close_date, next_activity_id, add_time, update_time"),
  },
  async (params) => {
    try {
      const queryParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );

      const response = await leadsV1Api.getLeads(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching leads:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching leads: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get archived leads
server.tool(
  "get-archived-leads",
  "Retrieve all archived leads that have been moved out of active pipeline. Shows historical lead data and conversion tracking for reporting purposes.",
  {
    limit: z.number().optional().describe("For pagination, the limit of entries to be returned. If not provided, 100 items will be returned."),
    start: z.number().optional().describe("For pagination, the position that represents the first result for the page"),
    owner_id: z.number().optional().describe("If supplied, only leads matching the given user will be returned. However, filter_id takes precedence over owner_id when supplied."),
    person_id: z.number().optional().describe("If supplied, only leads matching the given person will be returned. However, filter_id takes precedence over person_id when supplied."),
    organization_id: z.number().optional().describe("If supplied, only leads matching the given organization will be returned. However, filter_id takes precedence over organization_id when supplied."),
    filter_id: z.number().optional().describe("The ID of the filter to use"),
    sort: z.string().optional().describe("The field names and sorting mode separated by a comma (field_name_1 ASC, field_name_2 DESC). Supported fields: id, title, owner_id, creator_id, was_seen, expected_close_date, next_activity_id, add_time, update_time"),
  },
  async (params) => {
    try {
      const queryParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );

      const response = await leadsV1Api.getArchivedLeads(queryParams);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching archived leads:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching archived leads: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get lead by ID
server.tool(
  "get-lead",
  "Fetch detailed information about a single lead using its unique ID. Returns lead contact details, status, source, labels, and associated activities.",
  {
    leadId: z.string().describe("Lead ID (UUID format)"),
  },
  async ({ leadId }) => {
    try {
      const response = await leadsV1Api.getLead({ id: leadId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching lead ${leadId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching lead ${leadId}: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get lead permitted users
server.tool(
  "get-lead-permitted-users",
  "Retrieve all users who have access permissions to view and edit a specific lead. Shows user permissions and access levels for lead management.",
  {
    leadId: z.string().describe("Lead ID (UUID format)"),
  },
  async ({ leadId }) => {
    try {
      const response = await leadsV1Api.getLeadUsers({ id: leadId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error fetching permitted users for lead ${leadId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching permitted users for lead ${leadId}: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Create new lead
server.tool(
  "create-lead",
  "Add a new sales lead or prospect to Pipedrive. Create leads for potential customers before converting them to deals. Include contact details, source, and qualification status.",
  {
    title: z.string().describe("Lead title/name (required)"),
    owner_id: z.number().optional().describe("User ID of lead owner"),
    label_ids: z.array(z.string()).optional().describe("Array of label IDs to assign"),
    person_id: z.number().optional().describe("Person ID if lead is associated with existing contact"),
    organization_id: z.number().optional().describe("Organization ID if lead belongs to existing company"),
    source_name: z.string().optional().describe("Lead source name (e.g., 'Website', 'Cold Call')"),
    is_archived: z.boolean().optional().describe("Whether lead should be archived"),
    visible_to: z.enum(["1", "3", "5", "7"]).optional().describe("Visibility level"),
    value: z.object({
      amount: z.number().describe("Expected deal value"),
      currency: z.string().describe("Currency code (e.g., 'USD')"),
    }).optional().describe("Expected value of the lead"),
    was_seen: z.boolean().optional().describe("Whether lead has been seen/reviewed"),
  },
  async (params) => {
    try {
      const response = await leadsV1Api.addLead({ AddLeadRequest: params });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error creating lead:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error creating lead: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Update existing lead
server.tool(
  "update-lead",
  "Modify details of an existing lead such as title, owner, status, labels, or qualification information. Update lead progression and contact details.",
  {
    leadId: z.string().describe("Lead ID (UUID format)"),
    title: z.string().optional().describe("Lead title/name"),
    owner_id: z.number().optional().describe("User ID of lead owner"),
    label_ids: z.array(z.string()).optional().describe("Array of label IDs to assign"),
    person_id: z.number().optional().describe("Person ID if lead is associated with contact"),
    organization_id: z.number().optional().describe("Organization ID if lead belongs to company"),
    source_name: z.string().optional().describe("Lead source name"),
    is_archived: z.boolean().optional().describe("Whether lead should be archived"),
    visible_to: z.enum(["1", "3", "5", "7"]).optional().describe("Visibility level"),
    value: z.object({
      amount: z.number().describe("Expected deal value"),
      currency: z.string().describe("Currency code"),
    }).optional().describe("Expected value of the lead"),
    was_seen: z.boolean().optional().describe("Whether lead has been seen/reviewed"),
  },
  async ({ leadId, ...params }) => {
    try {
      const updateParams = Object.fromEntries(
        Object.entries(params).filter(([_, value]) => value !== undefined)
      );

      const response = await leadsV1Api.updateLead({
        id: leadId,
        UpdateLeadRequest: updateParams,
      });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error updating lead ${leadId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error updating lead ${leadId}: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Delete lead
server.tool(
  "delete-lead",
  "Remove a lead from Pipedrive permanently. Delete leads that are no longer viable prospects or have been converted to deals through other channels.",
  {
    leadId: z.string().describe("Lead ID (UUID format)"),
  },
  async ({ leadId }) => {
    try {
      const response = await leadsV1Api.deleteLead({ id: leadId });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error(`Error deleting lead ${leadId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error deleting lead ${leadId}: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === LEADS RELATED APIs ===

// Get lead labels
server.tool(
  "get-lead-labels",
  "Retrieve all lead labels and tags configured for lead categorization. Shows available labels for organizing and filtering sales prospects.",
  {},
  async () => {
    try {
      const response = await leadLabelsApi.getLeadLabels();
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching lead labels:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching lead labels: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Get lead sources
server.tool(
  "get-lead-sources",
  "Retrieve all lead source configurations for tracking where prospects originate. Shows lead sources like website, referral, cold call, social media, and marketing campaigns.",
  {},
  async () => {
    try {
      const response = await leadSourcesApi.getLeadSources();
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(cleanResponseData(response.data), null, 2),
          },
        ],
      };
    } catch (error) {
      console.error("Error fetching lead sources:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error fetching lead sources: ${getErrorMessage(error)}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// === PROMPTS ===

// Prompt for getting all deals
server.prompt("list-all-deals", "List all deals in Pipedrive", {}, () => ({
  messages: [
    {
      role: "user",
      content: {
        type: "text",
        text: "Please list all deals in my Pipedrive account, showing their title, value, status, and stage.",
      },
    },
  ],
}));

// Prompt for getting all persons
server.prompt("list-all-persons", "List all persons in Pipedrive", {}, () => ({
  messages: [
    {
      role: "user",
      content: {
        type: "text",
        text: "Please list all persons in my Pipedrive account, showing their name, email, phone, and organization.",
      },
    },
  ],
}));

// Prompt for getting all pipelines
server.prompt(
  "list-all-pipelines",
  "List all pipelines in Pipedrive",
  {},
  () => ({
    messages: [
      {
        role: "user",
        content: {
          type: "text",
          text: "Please list all pipelines in my Pipedrive account, showing their name and stages.",
        },
      },
    ],
  })
);

// Prompt for analyzing deals
server.prompt("analyze-deals", "Analyze deals by stage", {}, () => ({
  messages: [
    {
      role: "user",
      content: {
        type: "text",
        text: "Please analyze the deals in my Pipedrive account, grouping them by stage and providing total value for each stage.",
      },
    },
  ],
}));

// Prompt for analyzing contacts
server.prompt(
  "analyze-contacts",
  "Analyze contacts by organization",
  {},
  () => ({
    messages: [
      {
        role: "user",
        content: {
          type: "text",
          text: "Please analyze the persons in my Pipedrive account, grouping them by organization and providing a count for each organization.",
        },
      },
    ],
  })
);

// Prompt for analyzing leads
server.prompt("analyze-leads", "Analyze leads by status", {}, () => ({
  messages: [
    {
      role: "user",
      content: {
        type: "text",
        text: "Please search for all leads in my Pipedrive account and group them by status.",
      },
    },
  ],
}));

// Prompt for pipeline comparison
server.prompt(
  "compare-pipelines",
  "Compare different pipelines and their stages",
  {},
  () => ({
    messages: [
      {
        role: "user",
        content: {
          type: "text",
          text: "Please list all pipelines in my Pipedrive account and compare them by showing the stages in each pipeline.",
        },
      },
    ],
  })
);

// Prompt for finding high-value deals
server.prompt("find-high-value-deals", "Find high-value deals", {}, () => ({
  messages: [
    {
      role: "user",
      content: {
        type: "text",
        text: "Please identify the highest value deals in my Pipedrive account and provide information about which stage they're in and which person or organization they're associated with.",
      },
    },
  ],
}));

// Start the server with stdio transport
const transport = new StdioServerTransport();
server.connect(transport).catch((err) => {
  console.error("Failed to start MCP server:", err);
  process.exit(1);
});

console.error("Pipedrive MCP Server started");
