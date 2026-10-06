# AGENTS.md

This file defines the engineering principles and coding practices that must be followed when working in this repository.

These instructions apply to all coding agents and automated code changes.

When existing repository conventions conflict with these guidelines, inspect the surrounding code first and preserve consistency unless there is a strong reason to improve it.

---

# General Engineering Principles

Prefer simple, explicit, maintainable solutions.

Follow:

* KISS — Keep It Simple
* YAGNI — You Aren't Gonna Need It
* DRY where duplication represents the same knowledge
* high cohesion
* low coupling
* composition over inheritance
* explicit dependencies
* dependency inversion where it provides real value
* small, focused classes and methods
* clear ownership of state
* deterministic behavior

Do not introduce abstractions merely because an abstraction is possible.

Every interface, service, factory, wrapper, or architectural layer should have a concrete reason to exist.

Avoid speculative generalization.

Do not build infrastructure for hypothetical future requirements.

---

# Prefer Readability Over Cleverness

Code should be easy to understand by another experienced .NET developer without requiring unnecessary mental effort.

Prefer:

```csharp
var application = await applicationService.GetAsync(id, cancellationToken);
```

over unnecessarily abstract or clever implementations.

Use modern C# features when they improve clarity.

Do not use new language features merely because they are new.

Prefer explicit domain terminology over generic names such as:

```text
Manager
Helper
Processor
Utility
Handler
Thing
```

unless those names accurately represent the responsibility.

---

# Functional Programming Style

Use functional programming techniques where they improve correctness and readability.

The preferred library is:

```text
CSharpFunctionalExtensions
```

by Vladimir Khorikov.

Prefer the library's established abstractions rather than creating custom equivalents.

Important types include:

```csharp
Result
Result<T>
Result<T, E>
UnitResult<E>
Maybe<T>
```

Use functional composition where appropriate.

Useful operations include:

```csharp
Map
Bind
Ensure
Tap
MapError
Match
Finally
```

Prefer a pipeline such as:

```csharp
return FindApplication(id)
    .Ensure(
        application => application.IsEnabled,
        "Application is disabled.")
    .Bind(UpdateApplication)
    .Tap(LogUpdate);
```

over deeply nested imperative control flow.

---

# Result Instead of Expected Exceptions

Use `Result` for expected failures.

Examples of expected failures:

* validation failures
* entity not found
* invalid user input
* unsupported state
* authorization failure returned from an application operation
* Microsoft Graph operation rejected for a known reason
* requested resource already exists
* business-rule violation

Prefer:

```csharp
Result<Application> GetApplication(...)
```

instead of:

```csharp
Application GetApplication(...)
```

that throws an exception when an application simply does not exist.

Exceptions are for exceptional or unexpected failures.

Examples:

* programming errors
* corrupted application state
* infrastructure failures that cannot meaningfully be handled at the current level
* violated internal invariants
* impossible states

Do not use exceptions as normal control flow.

---

# Maybe for Optional Values

Use:

```csharp
Maybe<T>
```

when absence is a normal and meaningful outcome.

Prefer:

```csharp
Maybe<Application> FindApplication(...)
```

instead of returning `null` where practical.

Avoid unnecessary nullable-reference-state propagation when `Maybe<T>` expresses the intent better.

Do not wrap every nullable value in `Maybe<T>` mechanically.

Use it where absence is part of the domain/API semantics.

---

# Functional Core, Imperative Shell

Prefer keeping domain decisions and transformations as pure as practical.

Side effects should be pushed toward application boundaries.

For example:

```text
UI
 ↓
Application service
 ↓
Pure validation / transformations
 ↓
Microsoft Graph / persistence / external services
```

Avoid mixing:

```text
validation
remote calls
UI state
logging
serialization
domain rules
```

inside one method.

---

# Avoid Excessive Functional Ceremony

Functional programming should make the code simpler, not harder to understand.

Do not create long chains purely for style.

If this:

```csharp
if (application is null)
{
    return Result.Failure("Application was not found.");
}
```

is clearer than an elaborate functional expression, use the clearer code.

Functional programming is a tool, not a goal.

---

# Architecture

Prefer clear dependency direction.

High-level application logic should not depend directly on low-level implementation details where an abstraction provides useful separation.

Typical direction:

```text
UI
 ↓
Application services
 ↓
Domain / application logic
 ↓
Microsoft Graph SDK / external infrastructure
```

Keep Blazor components focused on presentation.

Blazor components should primarily manage:

* presentation
* UI state
* user interaction
* invoking application services

Do not place substantial Microsoft Graph or application logic directly in UI components.

---

# Composition Over Inheritance

Prefer composition unless inheritance models a genuine "is-a" relationship and provides clear value.

Avoid deep inheritance hierarchies.

Prefer:

```csharp
public sealed class ClaimsMappingPolicyService(
    GraphServiceClient graphClient,
    ILogger<ClaimsMappingPolicyService> logger)
```

over large base-service hierarchies.

Do not create base classes solely to share a small amount of code.

---

# Dependency Injection

Use constructor injection.

Dependencies must be explicit.

Prefer:

```csharp
public sealed class ApplicationService(
    GraphServiceClient graphClient,
    ILogger<ApplicationService> logger)
```

Avoid:

* service locator
* static service access
* hidden global dependencies
* manually resolving services from `IServiceProvider` without a concrete reason

Use appropriate lifetimes.

Be especially careful in Blazor Interactive Server with:

```text
Singleton
Scoped
Transient
```

Never store user-specific or tenant-specific state in singleton services.

---

# State Management

Avoid global mutable state.

Do not use static mutable fields for:

* current tenant
* current user
* access tokens
* Microsoft Graph clients
* configuration
* test state
* cached application objects

State belonging to a Blazor circuit/user must have an appropriate scoped lifetime.

Tenant-specific state must never leak between tenants.

---

# Async Programming

Use async/await end-to-end.

Never use:

```csharp
.Result
.Wait()
.GetAwaiter().GetResult()
```

unless there is an exceptional and well-documented reason.

Prefer:

```csharp
await operation.ExecuteAsync(cancellationToken);
```

All I/O operations should normally be asynchronous.

Pass `CancellationToken` through asynchronous call chains where appropriate.

For example:

```csharp
Task<Result<Application>> GetApplicationAsync(
    Guid id,
    CancellationToken cancellationToken = default);
```

Do not use:

```csharp
async void
```

except for APIs that explicitly require event-handler semantics.

---

# CancellationToken

Accept and propagate `CancellationToken` for:

* Microsoft Graph operations
* HTTP operations where custom HTTP is genuinely necessary
* long-running work
* database operations
* external service calls

Do not create arbitrary new `CancellationTokenSource` instances when the caller already provides a token.

---

# Microsoft Graph

Use the **official Microsoft Graph .NET SDK** as the primary and default way to access Microsoft Graph.

Prefer:

```text
Microsoft.Graph
GraphServiceClient
```

and the strongly typed request builders and models provided by the official SDK.

Before creating any custom Graph abstraction or HTTP implementation:

1. check whether the official Microsoft Graph .NET SDK already supports the required operation;
2. inspect the current SDK API;
3. use the SDK implementation when it exists.

Do **not** create custom HTTP services, REST wrappers, endpoint clients, request models, or response models for Microsoft Graph operations already supported by the official Microsoft Graph .NET SDK.

For example, prefer:

```csharp
await graphServiceClient
    .Applications[applicationId]
    .GetAsync(
        requestConfiguration =>
        {
            requestConfiguration.QueryParameters.Select =
            [
                "id",
                "appId",
                "displayName"
            ];
        },
        cancellationToken);
```

over manually implementing:

```text
GET https://graph.microsoft.com/v1.0/applications/{id}
```

with `HttpClient`.

The official SDK should also be preferred for:

* applications
* service principals
* users
* organization information
* policies
* directory roles
* paging
* `$ref` relationships
* create/update/delete operations

when supported by the installed SDK version.

Do not duplicate functionality already provided by the SDK.

---

# Raw Microsoft Graph HTTP Calls

Direct Microsoft Graph HTTP calls are an **exception**, not a normal design choice.

Use raw HTTP only when all of the following are true:

1. the required Microsoft Graph endpoint or operation is not supported by the current official Microsoft Graph .NET SDK;
2. using the SDK's supported extensibility mechanism cannot cleanly perform the operation;
3. the limitation has been verified against the current SDK;
4. the reason is documented in the code.

Do not choose raw HTTP merely because the SDK syntax is less familiar or slightly more verbose.

If raw Graph access is genuinely required, isolate it narrowly rather than creating a parallel Graph client architecture.

Do not create a generic custom REST abstraction around Microsoft Graph.

When the official SDK gains support for such an operation, prefer migrating to the SDK.

---

# Microsoft Graph API Version

Prefer Microsoft Graph:

```text
v1.0
```

Do not use:

```text
/beta
```

unless the required functionality genuinely does not exist in `v1.0`.

If `/beta` is required:

* isolate its usage
* document why it is necessary
* verify that there is no `v1.0` equivalent
* do not silently use it

---

# Microsoft Graph Queries

Use Graph query capabilities when appropriate:

```text
$select
$filter
$top
$orderby
$expand
```

Use the official SDK request configuration API rather than manually constructing URLs whenever possible.

Fetch only required properties.

Do not retrieve entire directory objects when only a few fields are needed.

Use paging through the mechanisms provided by the Graph SDK.

Do not load entire tenant collections into memory unnecessarily.

---

# Error Handling

Handle errors at the correct architectural level.

Infrastructure code should preserve enough information for higher layers to make useful decisions.

Prefer typed/structured errors over arbitrary strings for complex operations.

For example:

```csharp
public sealed record GraphError(
    int StatusCode,
    string? Code,
    string Message,
    string? RequestId);
```

Then:

```csharp
Result<Application, GraphError>
```

may be preferable to:

```csharp
Result<Application>
```

when callers need to distinguish error categories.

Translate Microsoft Graph SDK exceptions into application-level errors at an appropriate boundary when callers need to handle expected failures.

Do not catch:

```csharp
catch (Exception)
```

just to return a generic failure.

Only catch exceptions when:

* adding meaningful context
* translating an infrastructure exception into an application-level error
* performing required cleanup
* implementing an intentional boundary

Never silently swallow exceptions.

---

# Validation

Validate input as close as practical to the relevant application boundary.

Prefer explicit invariants.

Avoid passing invalid objects deep into the system and validating them later.

Use `Ensure` from CSharpFunctionalExtensions where it improves readability.

Example:

```csharp
return Result.Success(request)
    .Ensure(
        x => x.TenantId != Guid.Empty,
        "Tenant ID is required.")
    .Ensure(
        x => !string.IsNullOrWhiteSpace(x.DisplayName),
        "Display name is required.");
```

---

# Nullability

Enable nullable reference types.

Treat nullable warnings seriously.

Do not suppress nullable warnings with:

```csharp
!
```

unless correctness is already guaranteed and the reason is obvious.

Prefer modeling absence intentionally.

Use:

```csharp
Maybe<T>
```

where domain semantics justify it.

---

# Immutability

Prefer immutable data models where practical.

Use:

```csharp
record
```

and:

```csharp
init
```

when they accurately model the object.

Avoid mutable DTOs unless required by:

* serialization
* framework binding
* generated Microsoft Graph models
* UI editing

Do not mutate shared state unnecessarily.

---

# Collections

Expose the narrowest useful collection abstraction.

Prefer:

```csharp
IReadOnlyCollection<T>
IReadOnlyList<T>
IEnumerable<T>
```

when callers should not mutate the collection.

Avoid exposing mutable internal collections.

---

# LINQ

Use LINQ when it improves readability.

Avoid overly complicated LINQ expressions.

Do not hide expensive operations behind innocent-looking LINQ chains.

Prefer a normal loop when:

* logic has multiple branches
* debugging would be substantially easier
* performance characteristics need to be explicit
* the LINQ expression becomes difficult to understand

---

# Configuration

Use strongly typed configuration.

Prefer:

```csharp
IOptions<T>
IOptionsSnapshot<T>
```

where appropriate.

Validate configuration at startup.

Use:

```csharp
ValidateDataAnnotations()
Validate(...)
ValidateOnStart()
```

where useful.

Do not repeatedly read environment variables throughout application code.

Read external configuration at the composition boundary and expose it through strongly typed options/services.

---

# Secrets

Never commit:

* client secrets
* access tokens
* refresh tokens
* passwords
* certificates/private keys
* connection strings containing credentials

For local development prefer:

```text
dotnet user-secrets
```

Use secure secret providers in deployed environments.

Never log secrets.

Do not assume automated log masking will protect accidentally logged structured secrets.

---

# Logging

Use structured logging.

Prefer:

```csharp
logger.LogInformation(
    "Claims mapping policy {PolicyId} assigned to service principal {ServicePrincipalId}",
    policyId,
    servicePrincipalId);
```

instead of:

```csharp
logger.LogInformation(
    $"Claims mapping policy {policyId} assigned to {servicePrincipalId}");
```

Do not log:

* access tokens
* refresh tokens
* JWTs
* passwords
* secrets
* Authorization headers

Log identifiers and metadata that help troubleshooting.

Avoid excessive `Information` logging from frameworks.

Application logs may remain at `Information`, while framework categories should normally be `Warning` or above.

---

# Security

Security-sensitive operations must be explicit.

Never trust:

* route parameters
* browser-supplied tenant IDs
* object IDs
* client IDs
* policy IDs

without validation.

Authorization must be enforced server-side.

UI visibility is not authorization.

Do not expose Graph access tokens to client-side JavaScript.

Do not store Graph tokens in browser storage.

---

# Blazor

Use ASP.NET Core Blazor Web App with Interactive Server according to the existing project architecture.

## Use Code-Behind

Blazor components must use **code-behind files** for C# logic.

For a component:

```text
Applications.razor
```

place its C# implementation in:

```text
Applications.razor.cs
```

using a partial class:

```csharp
public partial class Applications
{
}
```

Keep `.razor` files focused on markup and presentation.

Do not place substantial C# code in:

```razor
@code {
    ...
}
```

Avoid `@code` blocks except for extremely small presentation-only cases where creating code-behind would add no value.

As a default rule:

```text
.razor
    → markup and presentation

.razor.cs
    → component state, lifecycle methods, event handlers,
      dependency access, orchestration, and C# logic
```

Prefer code-behind for:

* injected dependencies
* fields
* properties
* component lifecycle methods
* event handlers
* commands/actions
* loading state
* error state
* calls to application services
* navigation logic
* UI orchestration

Keep substantial domain and infrastructure logic out of both `.razor` and `.razor.cs`.

That logic belongs in application/services classes.

---

# Blazor Component Responsibilities

Blazor components should primarily handle:

* presentation
* UI state
* user interaction
* invoking application services
* displaying results

Do not place Microsoft Graph SDK operations directly inside Blazor components unless there is an exceptionally strong reason.

Prefer:

```text
Component code-behind
    ↓
Application/service layer
    ↓
GraphServiceClient
```

Keep components focused and testable.

Break large components into smaller components when they have distinct responsibilities.

Take Blazor circuit lifetime into account when managing user-specific state.

---

# Testing

Tests must be:

* deterministic
* independent
* readable
* repeatable
* focused on behavior
* valuable enough to justify their maintenance cost

Follow Arrange / Act / Assert where appropriate.

Test names should clearly describe the behavior or invariant being protected.

Prefer testing observable behavior rather than implementation details.

A unit test should normally have **one primary reason to fail**.

This does not mean mechanically enforcing one assertion per test.
Multiple assertions are appropriate when they collectively prove one
behavior, outcome, or invariant.

Use the **minimum number of assertions necessary** to establish the
behavior under test.

Split a test when its assertions verify independent behaviors that could
reasonably fail for unrelated reasons.

Prefer tests that remain valid after internal refactoring when externally
observable behavior has not changed.

Do not write tests merely to increase coverage.

Do not maximize the number of tests.
Maximize confidence provided by the test suite while minimizing its
long-term maintenance cost.

Before adding a test, consider:

> What realistic regression would this test detect?

If there is no meaningful answer, reconsider whether the test provides
enough value to justify its existence.

Avoid tests for trivial framework behavior, simple property accessors,
or implementation details unless they protect an important contract.

---

# Test Isolation

Avoid process-global mutable test state.

Be especially careful with:

```csharp
Environment.SetEnvironmentVariable(...)
```

because environment variables are process-wide.

Tests modifying:

* environment variables
* static state
* shared files
* fixed network ports
* external test tenants

must be explicitly isolated.

Do not assume disabling test-method parallelism solves parallelism between separate test assemblies or processes.

Tests sharing external resources must be designed with concurrency in mind.

---

# Unit Tests vs Integration Tests

Unit tests must not depend on:

* Microsoft Entra
* Microsoft Graph
* Internet connectivity
* external databases
* real credentials

Use test doubles for external dependencies.

Integration tests involving real Microsoft Entra or Graph must be clearly separated.

Do not make normal builds dependent on external integration-test infrastructure.

---

# Test Libraries

Follow the existing repository testing stack.

Where already used, prefer:

```text
xUnit v3
FluentAssertions
NSubstitute
```

Do not introduce another mocking/assertion framework without a concrete reason.

---

# Test Assertions

Prefer expressive assertions that communicate the behavioral expectation.

Use the minimum number of assertions necessary to prove the behavior
under test.

Do not apply a mechanical "one assertion per test" rule.

Several assertions are appropriate when they describe different aspects
of one logical result or invariant.

For example, when verifying that Microsoft Entra identifiers remain
semantically distinct, it is reasonable for one test to verify:

* service principal object ID
* application object ID
* application/client ID

when preserving those distinctions is the single behavior being tested.

Split the test when assertions verify independent concerns.

Avoid assertions against implementation details such as:

* exact internal method invocation counts
* exact invocation order
* private implementation state
* intermediate objects
* exact number of external requests

unless that detail is itself an intentional behavioral, performance,
or protocol contract.

Prefer state and output verification over interaction verification.

For `Result`, verify the success/failure state and the meaningful
returned value or error when both are necessary to establish the
behavior.

Avoid assertions that obscure the behavior under test.

---

# AI-Generated Tests

Do not generate tests mechanically for every method, branch, property,
or DTO.

Before creating a test:

1. identify the behavior or invariant being protected;
2. identify a realistic regression the test would catch;
3. check whether an existing test already protects that behavior;
4. choose the smallest useful test boundary;
5. avoid introducing test abstractions unless they reduce meaningful
   duplication or improve readability.

Do not create mocks merely because a dependency can be mocked.

Prefer real domain objects and simple deterministic collaborators.
Mock or fake external boundaries such as Microsoft Graph when needed.

Do not reproduce the production algorithm inside the test.

Do not make a test more complicated than the production behavior it
protects.

When modifying production code, prefer adding or changing the smallest
set of tests necessary to protect the changed behavior.

---

# Comments

Code should normally explain itself.

Comments should explain:

* why something unusual is necessary
* external constraints
* non-obvious behavior
* security considerations
* workarounds

Do not write comments that simply repeat the code.

Bad:

```csharp
// Get the application.
var application = GetApplication();
```

Useful:

```csharp
// Service principal object IDs are tenant-specific, so the object
// must be resolved again after switching tenants.
```

---

# XML Documentation

Do not add XML documentation mechanically to every internal type.

Add it where it provides useful API/context information.

Public reusable APIs should have documentation when their purpose or behavior is not obvious.

---

# Naming

Use domain-specific names.

Prefer:

```text
ClaimsMappingPolicy
ServicePrincipal
TenantContext
GraphOperation
ApplicationRegistration
```

over vague terms.

Methods that perform asynchronous work should end in:

```text
Async
```

unless framework conventions dictate otherwise.

Boolean names should read naturally:

```csharp
IsEnabled
HasPermission
CanEdit
ShouldRefresh
```

---

# Method Size and Complexity

Keep methods focused on one conceptual responsibility.

Do not enforce arbitrary line-count limits.

Refactor when:

* cyclomatic complexity grows unnecessarily
* nested conditions obscure the main flow
* a method performs unrelated responsibilities
* testing individual behavior becomes difficult

Prefer guard clauses.

Example:

```csharp
if (tenant is null)
{
    return Result.Failure("Tenant was not found.");
}

if (!tenant.IsEnabled)
{
    return Result.Failure("Tenant is disabled.");
}
```

over deeply nested conditions.

---

# Avoid Premature Optimization

Write correct and clear code first.

Optimize when:

* measurements show a problem
* scale characteristics clearly require it
* unnecessary remote/network/database work can obviously be avoided

Avoid unnecessary Microsoft Graph calls because remote operations are materially more expensive than in-process operations.

---

# Performance

For Microsoft Graph:

* fetch only required fields
* use `$select`
* use server-side filtering where supported
* page results
* avoid N+1 Graph calls
* use appropriate caching when safe
* make caches tenant-aware
* debounce interactive searches

Never trade correctness or tenant isolation for caching.

---

# Dependency Management

Reuse packages already present in the repository when suitable.

Before adding a dependency:

1. determine whether the .NET framework already provides the functionality;
2. determine whether an existing dependency already solves it;
3. verify that the dependency is actively maintained;
4. prefer stable releases;
5. avoid adding packages for trivial functionality.

For Microsoft Graph functionality, check the official Microsoft Graph .NET SDK before adding another Graph-related package or creating custom REST infrastructure.

Do not update unrelated package versions while implementing an unrelated feature.

---

# Change Discipline

Before making changes:

1. inspect relevant existing code;
2. understand existing architecture;
3. identify dependencies and callers;
4. check whether the Microsoft Graph SDK already supports required Graph operations;
5. make the smallest coherent change.

After changing code:

1. build;
2. run relevant tests;
3. inspect warnings;
4. fix introduced issues;
5. review the diff for unrelated changes.

Do not perform broad refactoring while implementing a small feature unless the refactoring is necessary.

---

# Preserve Existing Behavior

Do not change public behavior accidentally.

When modifying existing code:

* understand current callers;
* preserve backward compatibility unless the task explicitly requires a breaking change;
* add tests around behavior being changed;
* avoid unrelated formatting changes.

---

# No Blind Code Generation

Do not generate large amounts of boilerplate before inspecting the existing repository.

Do not assume architecture based only on the task description.

Inspect actual code first.

Prefer adapting to the existing codebase over replacing it.

Do not generate custom Microsoft Graph HTTP infrastructure before verifying that the official SDK lacks the required operation.

---

# Warnings

Treat compiler and analyzer warnings seriously.

Do not suppress warnings globally merely to make the build green.

Understand the warning first.

Suppress a warning only when:

* it is understood;
* it is demonstrably safe;
* suppression is narrowly scoped;
* the reason is documented where non-obvious.

---

# Build Verification

After a meaningful change, run the relevant build and tests.

At minimum before declaring a task complete:

```text
dotnet restore
dotnet build --configuration Release
dotnet test --configuration Release
```

Adapt commands to the existing repository structure.

Do not claim completion when the project does not compile.

---

# Agent Behavior

Before writing code:

1. inspect the relevant project and files;
2. understand the existing implementation;
3. identify the smallest appropriate change;
4. reuse existing abstractions;
5. check current official documentation when working with APIs that may have changed;
6. check the official Microsoft Graph .NET SDK before implementing Graph HTTP calls.

While implementing:

* work incrementally;
* compile frequently;
* run focused tests;
* preserve existing behavior;
* avoid unrelated changes;
* keep Blazor markup and component C# logic separated through code-behind.

Before finishing:

* review the complete diff;
* remove temporary debugging code;
* remove unused code;
* remove unused dependencies;
* ensure no secrets were introduced;
* run relevant tests;
* report important assumptions and limitations.

Do not leave commented-out experiments or temporary diagnostic code.

---

# Decision Priority

When multiple solutions are possible, prefer them in roughly this order:

1. correctness
2. security
3. simplicity
4. readability
5. maintainability
6. testability
7. performance
8. extensibility

Do not sacrifice the first six merely for speculative future extensibility.

---

# Core Philosophy

The desired codebase should be:

```text
Simple
Explicit
Composable
Testable
Secure
Predictable
Maintainable
```

Use functional techniques to make failure and optionality explicit.

Use object-oriented techniques where they model responsibilities naturally.

Do not treat functional programming or object-oriented programming as ideology.

Choose the technique that makes the domain and behavior clearest.

For Blazor:

```text
Keep markup in .razor.
Keep component C# logic in .razor.cs.
Keep business and infrastructure logic in services.
```

For Microsoft Graph:

```text
Use the official Microsoft Graph .NET SDK first.
Do not recreate APIs that the SDK already implements.
Use raw HTTP only as a verified exception.
```

When in doubt:

```text
Make invalid states difficult to represent.
Make expected failures explicit.
Keep side effects at the boundaries.
Prefer composition.
Keep dependencies visible.
Reuse official SDK functionality.
Keep UI markup separate from component logic.
Keep the solution as simple as the problem allows.
```

# Code formatting

Use .editorconfig to enforce consistent code formatting across the solution.