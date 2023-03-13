import { useTeams } from "msteams-react-base-component";
import * as React from "react";

export function UserProfile() {

  const [ name, setName ] = React.useState<string | undefined | null>();
  let url = new URL(window.location.href);

  React.useEffect(() => {
    setName(url.searchParams.get('name'));
  }, [url.searchParams]);

  return (
    <div>
      <h3>Profilo Utente (WIP)</h3>

      <div>
        <span>Nome: </span>
        <span>{name}</span>
      </div>

      <div>
        <h3>Url Parameters</h3>

        <ul>
          {
            Array.from(url.searchParams.entries())
              .map( ([name,value]) => 
                <li key={`${name}-${value}`}><b>{name}</b>: {value}</li>
              )
          }
        </ul>
      </div>
    </div>
  )
}