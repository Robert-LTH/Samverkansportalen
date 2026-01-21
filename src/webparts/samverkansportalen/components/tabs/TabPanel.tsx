import * as React from 'react';
import type { MainTabKey } from '../types';

interface ITabPanelProps {
  tabKey: MainTabKey;
  selectedKey: MainTabKey;
  children: React.ReactNode;
}

const TabPanel: React.FC<ITabPanelProps> = ({ tabKey, selectedKey, children }) => {
  const isSelected: boolean = tabKey === selectedKey;

  return (
    <div
      role="tabpanel"
      aria-hidden={!isSelected}
      hidden={!isSelected}
      className="samverkansportalen-tab-panel"
    >
      {isSelected ? children : null}
    </div>
  );
};

export default TabPanel;
