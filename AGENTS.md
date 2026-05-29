# Copilot Agents for Genesis Project

This project includes multiple specialized agents for SAP ABAP development and related technologies.

## Available Agents

### 1. **sap-abap-accelerator**
Primary agent for ABAP development with full MCP tool access.

**When to use:**
- Writing ABAP programs, reports, and modules
- Creating function modules and classes
- Implementing business logic
- Code refactoring and optimization

**Capabilities:**
- Full ABAP syntax support
- Standard library integration
- Dictionary object creation
- Transport management

**Example:**
```
/delegate "Create a customer master data report with ALV output"
```

### 2. **sap-odata-builder**
Specialized agent for OData and REST service development.

**When to use:**
- Creating OData services
- Building REST APIs for SAP integration
- Gateway service configuration
- Integration scenarios

### 3. **sap-fiori-developer**
UI5/Fiori application development agent.

**When to use:**
- Building Fiori applications
- UI5 component development
- Responsive design implementation
- User experience optimization

### 4. **sap-btp-cloud**
Cloud development agent for SAP Business Technology Platform.

**When to use:**
- CAP (Cloud Application Programming) development
- Microservices implementation
- BTP service integration
- Cloud security and deployment

### 5. **sap-transport-manager**
Change management and transport request handler.

**When to use:**
- Managing transport requests
- Coordinating releases
- Change management workflows
- Version control for SAP objects

### 6. **sap-test-automation**
Quality assurance and testing agent.

**When to use:**
- Creating ABAP Unit tests
- Integration testing
- Test data management
- Automation testing

## Skills

### Core Skills
- `abap-syntax-validator` - Validate ABAP code syntax
- `sap-dictionary-explorer` - Browse SAP dictionary
- `function-module-generator` - Generate FM templates
- `class-builder` - Create ABAP classes
- `report-designer` - Design ABAP reports

### Integration Skills
- `rfc-connector` - RFC function calls
- `odata-integrator` - OData consumption
- `idoc-processor` - IDoc handling
- `api-gateway` - API management

### Testing Skills
- `unit-test-runner` - Execute ABAP Unit tests
- `integration-tester` - Integration testing
- `performance-analyzer` - Code performance analysis
- `security-scanner` - Security vulnerability scanning

## MCP Servers

All configured MCP servers provide comprehensive tooling:

- **sap-devs**: Official SAP development CLI
- **sap-abap-adt**: ABAP Development Tools interface
- **sap-analytics**: Analytics and reporting
- **sap-security**: Security and authorization

## Usage Examples

### Basic ABAP Development
```
I need to create a function module that validates customer data.
Implement proper error handling and exception management.
```

### OData Service
```
Create an OData service for accessing sales orders with filtering
and paging support.
```

### Fiori Application
```
Build a Fiori app to display customer master data with edit capabilities.
Use ABAP backend as data source.
```

### Cloud Development
```
Develop a CAP service that exposes a custom business entity as OData.
```

## Configuration

All agents are configured in `.claude/CLAUDE.md` and use the setup from
`copilot-setup-steps.yml` workflow.

## Environment Setup

Before using agents, ensure:
1. SAP system connection is configured
2. ABAP development tools are installed
3. Required environment variables are set
4. Transport layer credentials are configured

## Further Reading

- [Copilot Agents Documentation](https://docs.github.com/copilot/concepts/agents)
- [SAP Development Documentation](https://help.sap.com)
- [ABAP Programming Guide](https://help.sap.com/doc/abapdocu_latest)
