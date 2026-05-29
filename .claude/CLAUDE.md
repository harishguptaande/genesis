# Copilot Configuration for SAP ABAP Development

This configuration enables GitHub Copilot with comprehensive SAP and ABAP development support.

## MCP Servers Configuration

```json
{
  "mcpServers": {
    "sap-devs": {
      "type": "stdio",
      "command": "sap-devs",
      "args": ["mcp", "serve"],
      "disabled": false
    },
    "sap-abap-adt": {
      "type": "stdio",
      "command": "npx",
      "args": ["@sap/adt-cli", "mcp"],
      "disabled": false
    }
  }
}
```

## Agent Instructions

### SAP ABAP Development
- Focus on ABAP syntax and best practices
- Use ADT (ABAP Development Tools) compatible code
- Follow SAP naming conventions (ZZ*, ZY* for custom objects)
- Ensure compatibility with SAP standard modules (MM, FI, SD, HR, etc.)

### Available Tools
- SAP ADT REST API access
- ABAP code analysis and refactoring
- Function module and class creation
- BDC (Batch Data Communication) automation
- OData service development
- RFC (Remote Function Call) integration

### Coding Standards
- Use ABAP Objects syntax
- Implement proper error handling with TRY-CATCH
- Follow SOLID principles
- Use type-safe declarations
- Implement logging and debugging support

## Skills Included

### 1. **sap-abap-accelerator**
   - Full-featured ABAP development assistant
   - Unrestricted MCP tool access
   - Best practices and code generation

### 2. **sap-odata-builder**
   - OData service creation and management
   - REST API design for SAP systems
   - Gateway service configuration

### 3. **sap-fiori-developer**
   - UI5/Fiori app development
   - Responsive design patterns
   - SAP Fiori Elements integration

### 4. **sap-btp-cloud**
   - SAP Business Technology Platform development
   - Cloud Application Programming (CAP)
   - Microservices architecture

### 5. **sap-transport-manager**
   - Transport request management
   - Change management
   - Release coordination

### 6. **sap-test-automation**
   - ABAP unit testing (ABAP Unit)
   - Integration testing
   - Test data management

## Custom Instructions

### When writing ABAP code:
1. Always specify package assignment (e.g., ZZ_MYAPP)
2. Include proper authorization checks
3. Use standardized include files for common functions
4. Implement logging using standard FM or CL_* classes
5. Follow dictionary object naming conventions

### When creating function modules:
1. Define clear interface parameters
2. Include comprehensive exception handling
3. Document parameters with descriptions
4. Mark remote-enabled if needed (RFC)
5. Use proper data types from ABAP dictionary

### For reports and programs:
1. Use REPORT or PROGRAM statement appropriately
2. Include selection screens for input
3. Use ALV (ABAP List Viewer) for output
4. Implement role-based authorization
5. Add help texts for UI elements

## Environment Variables (Set in GitHub)

```
SAP_HOST=your-sap-host
SAP_PORT=your-sap-port
SAP_CLIENT=your-sap-client
SAP_LANGUAGE=EN
ABAP_DEVELOPMENT_TOOLS=enabled
```

## Resources

- [SAP ABAP Naming Conventions](https://help.sap.com)
- [ABAP Development Tools Documentation](https://tools.hana.ondemand.com)
- [SAP CAP Documentation](https://cap.cloud.sap)
- [SAP Fiori Design Guidelines](https://experience.sap.com/fiori-design-web)
