import * as React from "react";
import Header from "./Header";
import LegalEditor from "./LegalEditor";
import { makeStyles } from "@fluentui/react-components";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    height: "100vh",
    display: "flex",
    flexDirection: "column",
    overflow: "hidden",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      <Header logo="assets/silks-logo.svg" title={props.title} />
      <LegalEditor />
    </div>
  );
};

export default App;
