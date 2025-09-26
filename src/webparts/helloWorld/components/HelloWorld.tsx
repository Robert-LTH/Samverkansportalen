import * as React from 'react';

import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';

const HelloWorld: React.FC<IHelloWorldProps> = ({ description }) => (
  <section className={styles.helloWorld}>
    <div className={styles.container}>
      <span className={styles.title}>Welcome to SharePoint!</span>
      <p className={styles.subTitle}>Customize your SharePoint experiences using web parts.</p>
      <p className={styles.description}>{description}</p>
    </div>
  </section>
);

export default HelloWorld;
