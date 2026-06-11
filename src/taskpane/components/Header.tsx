import * as React from "react";
import { Image, tokens, makeStyles } from "@fluentui/react-components";

export interface HeaderProps {
  title: string;
  logo: string;
}

const useStyles = makeStyles({
  welcome__header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-start",
    paddingTop: "12px",
    paddingBottom: "12px",
    paddingLeft: "16px",
    paddingRight: "16px",
    backgroundColor: "#111316",
  },
});

const Header: React.FC<HeaderProps> = (props: HeaderProps) => {
  const { title, logo } = props;
  const styles = useStyles();

  return (
    <section className={styles.welcome__header}>
      <Image height="30" src={logo} alt={title} />
    </section>
  );
};

export default Header;
