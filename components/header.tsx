
import classes from './header.module.css';


export default function Header() {
  return (
    <header className={classes.header}>
      <p role="heading" aria-level={1}>Hualpén · Control de peajes</p>
      <h2>Revisión de cobros de autopistas</h2>
      <p>
        Compara cada cobro con la tarifa esperada y calcula la rebaja sugerida
        para los 10 buses.
      </p>
    </header>
  );
}