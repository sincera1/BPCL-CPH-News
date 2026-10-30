import * as React from "react";
import Select, { CSSObjectWithLabel, SingleValue } from "react-select";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";

import ErrorService from "../services/ErrorService";
 
export interface ISPUser {
  Id: number;
  Title: string;
  Email: string;
  LoginName?: string; 
}
  
export interface ISPGroup {
  Id: number;
  Title: string;
} 
  
export type Option = { 
  label: string;
  value: number;
  type: "User" | "Group"; 
  title: string;
  email?: string;
  loginName?: string;
};
 
interface IPeoplePickerProps {
  context: unknown;
  selectedValue: Option | null;
  onChange: (user: Option | null) => void;
  label?: string;
  isClearable?: boolean; 
  placeholder?: string;
  minSearchLength?: number;
  debounceMs?: number;
}
 
const PeoplePicker: React.FC<IPeoplePickerProps> = ({
  context,
  
  selectedValue,
  onChange,
  label,
  placeholder = "Search people or groups...",
  minSearchLength = 2,
  debounceMs = 400,
}) => {
  const sp = React.useMemo(() => spfi().using(SPFx(context as any)), [context]);
 
  const [options, setOptions] = React.useState<Option[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
 
  // Debounce helper
  const debounce = <T extends (...args: any[]) => void>(fn: T, delay = 300) => {
    let timer: number;
    return (...args: Parameters<T>) => {
      window.clearTimeout(timer);
      timer = window.setTimeout(() => fn(...args), delay);
    };
  };
 
  const searchUsersAndGroups = async (term: string): Promise<void> => {
    if (!term || term.length < minSearchLength) {
      setOptions([]);
      return;
    }
 
    setLoading(true);
    try {
      const [users, groups]: [ISPUser[], ISPGroup[]] = await Promise.all([
        sp.web.siteUsers(),
        sp.web.siteGroups(),
      ]);
 
      const t = term.toLowerCase();
 
      const userOptions: Option[] = users
        .filter(
          (u) =>
            (u.Title || "").toLowerCase().indexOf(t) !== -1 ||
            (u.Email || "").toLowerCase().indexOf(t) !== -1
        )
        .map((u) => ({
          label: `${u.Title}${u.Email ? ` (${u.Email})` : ""}`,
          value: u.Id,
          type: "User",
          title: u.Title,
          email: u.Email,
          loginName: u.LoginName,
        }));
 
      const groupOptions: Option[] = groups
        .filter((g) => (g.Title || "").toLowerCase().indexOf(t) !== -1)
        .map((g) => ({
          label: `${g.Title} (Group)`,
          value: g.Id,
          type: "Group",
          title: g.Title,
        }));
 
      setOptions([...userOptions, ...groupOptions]);
    } catch (error) {

  ErrorService.logError(
    "PeoplePicker.searchUsersAndGroups",
    error
  ); 

  setOptions([]);
} finally {
      setLoading(false);
    }
  };
 
  const searchDebounced = React.useMemo(
    () => debounce(searchUsersAndGroups, debounceMs),
    [debounceMs]
  );
 
  return (
    <div className="mb-3">
      {label && <label className="form-label fw-semibold">{label}</label>}
 
      <Select
        isMulti={false} 
        isClearable={true}
        options={options}
        value={selectedValue}
        onChange={(value: SingleValue<Option>) => onChange(value ?? null)}
        onInputChange={(inputValue: string) => searchDebounced(inputValue)}
        placeholder={placeholder}
        isLoading={loading}
        noOptionsMessage={() => "Type to search users or groups..."}
        styles={{
          control: (base: CSSObjectWithLabel) => ({
            ...base,
            borderRadius: 6,
          }),
          menu: (base: CSSObjectWithLabel) => ({
            ...base,
            zIndex: 9999,
          }),
        }}
      />
    </div>
  );
};
 
export default PeoplePicker;