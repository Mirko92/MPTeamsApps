import { Button } from "@fluentui/react-northstar";
import * as React from "react";

export function LoginMSAL() {
  return (
    <div>
      <p>Effettua l'accesso</p>
      
      <div>
        <Button onClick={() => alert("It worked!")}>
          A sample button
        </Button>
      </div>
    </div>
  )
}