/**
 * Native-style group picker built on Zenith's GroupsFilter.
 *
 * Ported directly from advanced_report_builder/src/components/GroupFilterPicker.tsx
 * so EHZ's group selector behaves identically to ARB's: multi-select,
 * hierarchical, AND/OR filtering, surfaces flat IDs to the parent.
 */

import { useCallback, useMemo, useState } from "react";
import {
  GroupsFilter,
  RelationOperator,
  type IGroupItem,
  type IGroupsFilterTotalState,
  type IFilterState,
} from "@geotab/zenith";
import type { GeotabGroup } from "../types";

export interface GroupFilterPickerProps {
  /** Groups already fetched by App.tsx, keyed by id. */
  groupsById: Map<string, GeotabGroup>;
  /** Initial selection (defaults to Company group when omitted). */
  initialGroupIds?: string[];
  /** Called whenever the selection changes. */
  onChange: (groupIds: string[]) => void;
  /** Called when GroupsFilter surfaces an error. */
  onError?: (e: Error) => void;
}

function buildGroupItemArr(byId: Map<string, GeotabGroup>): IGroupItem[] {
  const items: IGroupItem[] = [];
  byId.forEach((g) => {
    items.push({
      id: g.id,
      name: g.name && g.name.length > 0 ? g.name : g.id,
      children:
        Array.isArray(g.children) && g.children.length > 0
          ? g.children.map((c) => ({ id: c.id }))
          : undefined,
    });
  });
  items.sort((a, b) => {
    if (a.id === "GroupCompanyId") return -1;
    if (b.id === "GroupCompanyId") return 1;
    return (a.name ?? "").localeCompare(b.name ?? "");
  });
  return items;
}

function flattenIds(state: IFilterState | undefined): string[] {
  if (!state || !Array.isArray(state.items)) return [];
  const out: string[] = [];
  for (const item of state.items) {
    if (
      item &&
      typeof item === "object" &&
      "items" in item &&
      Array.isArray((item as IFilterState).items)
    ) {
      out.push(...flattenIds(item as IFilterState));
    } else if (
      item &&
      typeof item === "object" &&
      "id" in item &&
      typeof (item as { id: string }).id === "string"
    ) {
      out.push((item as { id: string }).id);
    }
  }
  return out;
}

export function GroupFilterPicker({
  groupsById,
  initialGroupIds,
  onChange,
  onError,
}: GroupFilterPickerProps) {
  const [state, setState] = useState<IGroupsFilterTotalState>(() => {
    const ids =
      initialGroupIds && initialGroupIds.length > 0
        ? initialGroupIds
        : ["GroupCompanyId"];
    return {
      groups: {
        relation: RelationOperator.OR,
        items: ids.map((id) => ({ id })),
      },
      sideWide: false,
    };
  });

  const dataLoader = useCallback(async (): Promise<IGroupItem[]> => {
    return buildGroupItemArr(groupsById);
  }, [groupsById]);

  const handleChange = useCallback(
    (next: IGroupsFilterTotalState) => {
      setState(next);
      onChange(flattenIds(next.groups));
    },
    [onChange]
  );

  const errorHandler = useCallback(
    (e: Error) => {
      console.error("[EHZ] GroupsFilter error:", e);
      onError?.(e);
    },
    [onError]
  );

  const initialFilterState = useMemo(() => state, []); // eslint-disable-line react-hooks/exhaustive-deps

  return (
    <GroupsFilter
      dataLoader={dataLoader}
      onChange={handleChange}
      errorHandler={errorHandler}
      initialFilterState={initialFilterState}
    />
  );
}
