import * as React from 'react';
import styles from './SpfxSuraMymessages.module.scss';
import { ISpfxSuraMymessagesProps } from './ISpfxSuraMymessagesProps';
import { escape } from '@microsoft/sp-lodash-subset';

import {ISpfxSuraMymessagesState} from './ISpfxSuraMymessagesState';
import {MSGraphClient} from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import {
  Persona,
  PersonaSize
} from 'office-ui-fabric-react/lib/components/Persona';
import {Link} from 'office-ui-fabric-react/lib/components/Link';
import { IsFocusVisibleClassName } from '@uifabric/utilities/lib';

export default class SpfxSuraMymessages extends React.Component<ISpfxSuraMymessagesProps, ISpfxSuraMymessagesState> {
  constructor(props: ISpfxSuraMymessagesProps) {
    super(props);

    this.state = {
      name: '',
      email: '',
      phone: '',
      mails: [],
      events: []
    };
  }

  public componentDidMount(): void {
    
    this.props.graphClient.
    api('/me')
    .get((error: any, user: MicrosoftGraph.User, rawResponse?: any) => {
      this.setState({
        name: user.displayName,
        email: user.mail,
        phone: user.businessPhones[0],
      });
    });
    let correos = [];
    this.props.graphClient
    .api('me/messages')
    .select('subject,bodyPreview,sender')
    .get((error: any, messages: MicrosoftGraph.User, rawResponse?: any) => {
      console.log(messages)
      correos.push(messages);
      this.setState({
        mails: correos
      });
    });
    let eventos = []
    this.props.graphClient
    .api('me/events')
    .select('subject,bodyPreview,organizer,start,end,location')
    .get((error: any, messages: MicrosoftGraph.User, rawResponse?: any) => {
      console.log(messages)
      eventos.push(messages);
      this.setState({
        events: eventos
      });
    });
  }

  public render(): React.ReactElement<ISpfxSuraMymessagesProps> {
    
    return (
      <div className={ styles.spfxSuraMymessages }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to Digital Work Place SURA!</span>
              <p className={ styles.subTitle }>My Profile</p>
              <Persona primaryText={this.state.name}
                secondaryText={this.state.email}
                tertiaryText={this.state.phone}
                size={PersonaSize.size100}
                />
            </div>
          </div>
        </div>
      </div>
      
    );
  }
}
