# meta descriptions

Macro: [Skycode - meta descriptions.bas](<../../../vba scripts/Skycode - meta descriptions.bas>)

## Purpose

Sets each actor's `Meta.Description` to its configured description. If no configured value exists, it sanitises the current value.

## Host and entry point

- Composer VBA through an Access host. Run `SetMetaDescriptionfromDescription_cfg`.

## Inputs

- The active Composer document. All `CLitActorMesh` and `CLitActorView` actors are processed.
- `Meta.Description`, plus the configured `Meta.Description_cfg` (preferred) or `Meta.DESCRIPTION_cfg`.

## Outputs

- Writes `Meta.Description` through the Composer property APIs.
- Log: `C:\OneDrive - Skyfacture Pty Ltd\Skyfacture\3 - Economics\skycode\logs\meta descriptions.txt`.

## Side effects

- Changes document properties. The log is overwritten. Nothing is saved.

## Gotchas

- A description is valid only if it has no digit, bracket, slash, backslash, comma, non-ASCII character, and no dot except at the end. Invalid values are sanitised.
- Sanitising only normalises whitespace. Configured apostrophe entities are replaced with `/`.
