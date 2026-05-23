# Review Checklist Template

## Scope

- [ ] Change matches the task.
- [ ] No unrelated refactor.
- [ ] No accidental formatting noise.

## Backend

- [ ] Code compiles.
- [ ] Business rules are in the right layer.
- [ ] API status codes are correct.
- [ ] Validation is explicit.
- [ ] Authorization is enforced server-side.
- [ ] No EF entities are exposed as API DTOs.
- [ ] CancellationToken is passed where useful.

## Frontend

- [ ] TypeScript types are explicit.
- [ ] No unnecessary `any`.
- [ ] Component remains focused.
- [ ] API calls are in services/data-access.
- [ ] Loading/error/empty states are handled.
- [ ] Forms have validation messages.

## Security

- [ ] No secrets committed.
- [ ] No sensitive data logged.
- [ ] Access control is checked.
- [ ] Error messages do not leak internals.

## Quality

- [ ] Tests added or updated where useful.
- [ ] Build/test commands were run or limitation explained.
- [ ] `.ai/` memory was updated if needed.
