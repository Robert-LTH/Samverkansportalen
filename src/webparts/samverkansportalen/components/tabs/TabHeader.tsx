import * as React from 'react';
import { TabList, Tab } from '@fluentui/react-components';
import styles from '../Samverkansportalen.module.scss';
import type { MainTabKey } from '../types';

interface ITabHeaderItem {
  key: MainTabKey;
  label: string;
}

interface ITabHeaderProps {
  items: ITabHeaderItem[];
  selectedKey: MainTabKey;
  onSelect: (key: MainTabKey) => void;
}

// Work around React 17 typings vs Fluent UI v9 tab component generics in SPFx.
type FluentComponent = React.ComponentType<Record<string, unknown>>;
const FluentTabList = TabList as unknown as FluentComponent;
const FluentTab = Tab as unknown as FluentComponent;

const TabHeader: React.FC<ITabHeaderProps> = ({ items, selectedKey, onSelect }) => (
  <FluentTabList
    className={styles.pivotFloating}
    selectedValue={selectedKey}
    onTabSelect={(_event: React.SyntheticEvent, data: { value: unknown }) => {
      if (data.value === selectedKey) {
        return;
      }

      const nextValue: string = typeof data.value === 'string' ? data.value : String(data.value);
      onSelect(nextValue as MainTabKey);
    }}
    data-samverkansportalen-tablist="true"
  >
    {items.map((item) => (
      <FluentTab key={item.key} value={item.key}>
        {item.label}
      </FluentTab>
    ))}
  </FluentTabList>
);

export default TabHeader;
