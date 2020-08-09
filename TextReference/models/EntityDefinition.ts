import { UserLocalizedLabel } from "./UserLocalizedLabel";

export class EntityDefinition
{
    public LogicalName : string;
    public EntitySetName : string;
    public PrimaryNameAttribute : string;
    public PrimaryIdAttribute : string;
    public UserLocalizedLabel : UserLocalizedLabel;
    public Icon : string;
}