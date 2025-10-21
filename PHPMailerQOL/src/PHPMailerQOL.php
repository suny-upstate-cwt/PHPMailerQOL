<?php
/**
 * PHPMailerQOL - Enhance the PHPMailer class with quality-of-life
 *                method additions and overrides. The methods strive
 *                to be forgiving in regards to PHP error avoidance.
 *                All overrides strive to retain parent method behavior
 *                so this extension can be dropped into an existing
 *                codebase.
 *
 * @author    German (Gare-mon) Drulyk <drulykg@upstate.edu>
 * @copyright 2025 German (Gare-mon) Drulyk
 * @license   https://opensource.org/license/mit MIT License
 * @note      THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 *            EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
 *            OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 *            NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
 *            HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
 *            WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
 *            FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
 *            OTHER DEALINGS IN THE SOFTWARE.
 */

class PHPMailerQOL extends PHPMailer\PHPMailer\PHPMailer
{
    /**
     * Default Address Domain
     * Set a default domain to be appended to an addresss
     * Private - enforce setting via setter
     *
     * @var string
     */
    private $DefaultAddressDomain = '';
    
    /**
     * add* methods have been overridden to support address juggling if desired
     *
     * @var bool
     */
    public $JuggleAddressAdds = false;
    
    /**
     * Constructor. See parent PHPMailer class.
     *
     * @param bool $exceptions See parent PHPMailer class.
     *
     * @return void
     */
    public function __construct( $exceptions = null )
    {
        // Construct the parent class
        parent::__construct( $exceptions );
    }
    
    /**
     * Is HTML mode enabled?
     *
     * @return bool
     */
    public function getIsHTML()
    {
        return ( $this->ContentType === $this::CONTENT_TYPE_TEXT_HTML );
    }
    
    /**
     * Get default domain.
     *
     * @return string
     */
    public function getDefaultAddressDomain()
    {
        return $this->DefaultAddressDomain;
    }
    
    /**
     * Set default domain.
     *
     * @param string $domain
     *
     * @return string
     */
    public function setDefaultAddressDomain( $domain )
    {
        // Remove everything leading up to the final @ from $domain and trim() it
        // If there is no @ symbol, great! The whole string will be used
        $this->DefaultAddressDomain = ( is_scalar( $domain ) ? trim( preg_replace( '/\A.*@([^@]*)\Z/s', '$1', $domain ) ) : '' );
        
        return $this->DefaultAddressDomain;
    }
    
    /**
     * Append default domain to a given address if no domain is detected.
     *
     * @param string $address
     *
     * @return string
     */
    public function appendDefaultAddressDomain( $address )
    {
        static $trim_chars = "@ \n\r\t\v\x00";
        
        if(
            $this->DefaultAddressDomain !== '' &&
            is_scalar( $address ) &&
            trim( $address, $trim_chars ) !== '' &&
            strpos( trim( $address, $trim_chars ), '@' ) === false
        )
        {
            $address = trim( $address, $trim_chars ).'@'.$this->DefaultAddressDomain;
        }
        
        return $address;
    }
    
    /**
     * Override setFrom() to readily parse different input types.
     *
     * @param string|array $addresses
     * @param string|array $names
     * @param bool $auto
     *
     * @return bool
     */
    public function setFrom( $addresses, $names = [], $auto = true )
    {
        $return = false;
        
        foreach( $this->coalesceAddressesAndNames( $addresses, $names ) as $address => $name )
        {
            // Did the most recent setFrom() succeed?
            $return = parent::setFrom( $address, $name, $auto );
        }
        
        return $return;
    }
    
    /**
     * Clear TO addresses and add new ones.
     *
     * @param string|array $addresses
     * @param string|array $names
     *
     * @return bool
     */
    public function setAddress( $addresses, $names = [] )
    {
        $this->clearAddresses();
        return $this->addAddress( $addresses, $names );
    }
    
    /**
     * Alias for setAddress().
     *
     * @param string|array $addresses
     * @param string|array $names
     *
     * @return bool
     */
    public function setTO( $addresses, $names = [] )
    {
        return $this->setAddress( $addresses, $names );
    }
    
    /**
     * Clear CC addresses and add new ones.
     *
     * @param string|array $addresses
     * @param string|array $names
     *
     * @return bool
     */
    public function setCC( $addresses, $names = [] )
    {
        $this->clearCCs();
        return $this->addCC( $addresses, $names );
    }
    
    /**
     * Clear BCC addresses and add new ones.
     *
     * @param string|array $addresses
     * @param string|array $names
     *
     * @return bool
     */
    public function setBCC( $addresses, $names = [] )
    {
        $this->clearBCCs();
        return $this->addBCC( $addresses, $names );
    }
    
    /**
     * Clear Reply-Tos addresses and add new ones.
     *
     * @param string|array $addresses
     * @param string|array $names
     *
     * @return bool
     */
    public function setReplyTo( $addresses, $names = [] )
    {
        $this->clearReplyTos();
        return $this->addReplyTo( $addresses, $names );
    }
    
    /**
     * Override addAddress() to readily parse different input types.
     * Behavior ammended so that addresses can get juggled into position.
     *
     * @param string|array $addresses
     * @param string|array $names
     *
     * @return bool
     */
    public function addAddress( $addresses, $names = [] )
    {
        // Remove this address if it exists somewhere, even if already in TO
        // This allows for email address case and/or vanity name changes
        ( $this->JuggleAddressAdds ? $this->removeRecipient( $addresses ) : null );
        
        return $this->addOrEnqueueAnAddressList( 'to', $addresses, $names );
    }
    
    /**
     * Alias for addAddress().
     *
     * @param string|array $addresses
     * @param string|array $names
     *
     * @return bool
     */
    public function addTO( $addresses, $names = [] )
    {
        return $this->addAddress( $addresses, $names );
    }
    
    /**
     * Override addCC() to readily parse different input types.
     * Behavior ammended so that addresses can get juggled into position.
     *
     * @param string|array $addresses
     * @param string|array $names
     *
     * @return bool
     */
    public function addCC( $addresses, $names = [] )
    {
        // Remove this address if it exists somewhere, even if already in CC
        // This allows for email address case and/or vanity name changes
        ( $this->JuggleAddressAdds ? $this->removeRecipient( $addresses ) : null );
        
        return $this->addOrEnqueueAnAddressList( 'cc', $addresses, $names );
    }
    
    /**
     * Override addBCC() to readily parse different input types.
     * Behavior ammended so that addresses can get juggled into position.
     *
     * @param string|array $addresses
     * @param string|array $names
     *
     * @return bool
     */
    public function addBCC( $addresses, $names = [] )
    {
        // Remove this address if it exists somewhere, even if already in BCC
        // This allows for email address case and/or vanity name changes
        ( $this->JuggleAddressAdds ? $this->removeRecipient( $addresses ) : null );
        
        return $this->addOrEnqueueAnAddressList( 'bcc', $addresses, $names );
    }
    
    /**
     * Override addReplyTo() to readily parse different input types.
     * Behavior ammended so that addresses can get juggled into position.
     *
     * @param string|array $addresses
     * @param string|array $names
     *
     * @return bool
     */
    public function addReplyTo( $addresses, $names = [] )
    {
        // Remove this address if it exists as a Reply-To
        // This allows for email address case and/or vanity name changes
        ( $this->JuggleAddressAdds ? $this->removeReplyTo( $addresses ) : null );
        
        return $this->addOrEnqueueAnAddressList( 'Reply-To', $addresses, $names );
    }
    
    /**
     * Wrapper for PHPMailer->addOrEnqueueAnAddress() to readily parse different input types.
     *
     * @param string $kind
     * @param string|array $addresses
     * @param string|array $names
     *
     * @return bool
     */
    protected function addOrEnqueueAnAddressList( $kind, $addresses, $names = [] )
    {
        // Assume failure unless addOrEnqueueAnAddress() succeeds even once
        $return = false;
        
        foreach( $this->coalesceAddressesAndNames( $addresses, $names ) as $address => $name )
        {
            // A single success should cause $return to stay true
            $return = ( $this->addOrEnqueueAnAddress( $kind, $address, $name ) ? true : $return );
        }
        
        return $return;
    }
    
    /**
     * Override addOrEnqueueAnAddress() to resolve $kind
     *
     * @param string $kind
     * @param string $address
     * @param string $name
     *
     * @return bool
     */
    protected function addOrEnqueueAnAddress( $kind, $address, $name )
    {
        // Resolve $kind to a key position
        $kind = $this->resolveKindToKey( $kind );
        
        // PHPMailer->addOrEnqueueAnAddress() expects a hyphenated "Reply-To"
        $kind = ( $kind === 'ReplyTo' ? 'Reply-To' : $kind );
        
        return parent::addOrEnqueueAnAddress( $kind, $address, $name );
    }
    
    /**
     * Remove an address from the TO list.
     *
     * @param string|array $addresses
     *
     * @return bool
     */
    public function removeAddress( $addresses )
    {
        return $this->removeAnAddressList( 'to', $addresses );
    }
    
    /**
     * Alias for removeAddress().
     *
     * @param string|array $addresses
     *
     * @return bool
     */
    public function removeTO( $addresses )
    {
        return $this->removeAddress( $addresses );
    }
    
    /**
     * Remove an address from the CC list.
     *
     * @param string|array $addresses
     *
     * @return bool
     */
    public function removeCC( $addresses )
    {
        return $this->removeAnAddressList( 'cc', $addresses );
    }
    
    /**
     * Remove an address from the BCC list.
     *
     * @param string|array $addresses
     *
     * @return bool
     */
    public function removeBCC( $addresses )
    {
        return $this->removeAnAddressList( 'bcc', $addresses );
    }
    
    /**
     * Remove an address from any recipient list.
     *
     * @param string|array $addresses
     *
     * @return bool
     */
    public function removeRecipient( $addresses )
    {
        return (
            $this->removeTO( $addresses ) ||
            $this->removeCC( $addresses ) ||
            $this->removeBCC( $addresses )
        );
    }
    
    /**
     * Remove an address from the Reply-To list.
     *
     * @param string|array $addresses
     *
     * @return bool
     */
    public function removeReplyTo( $addresses )
    {
        return $this->removeAnAddressList( 'Reply-To', $addresses );
    }
    
    /**
     * Wrapper for removeAnAddress() to readily parse different input types.
     *
     * @param string $kind
     * @param string|array $addresses
     *
     * @return bool
     */
    protected function removeAnAddressList( $kind, $addresses )
    {
        // Assume failure unless removeAnAddress() succeeds even once
        $return = false;
        
        foreach( $this->coalesceAddressesAndNames( $addresses, [] ) as $address => $name )
        {
            // A single success should cause $return to stay true
            $return = ( $this->removeAnAddress( $kind, $address ) ? true : $return );
        }
        
        return $return;
    }
    
    /**
     * Remove an address from the type specified.
     *
     * @param string $kind
     * @param string|array $addresses
     *
     * @return bool
     */
    protected function removeAnAddress( $kind, $address )
    {
        // Assume failure unless $address is found in $kind position
        $return = false;
        
        // Resolve $kind to a key position
        $kind = $this->resolveKindToKey( $kind );
        
        if( $kind )
        {
            // Compare addresses case-insensitively
            $address = strtolower( trim( $address ) );
            
            // Loop the address storage key
            foreach( $this->{$kind} as $k => $email )
            {
                // The emails are in position zero
                $email = strtolower( trim( $email[ 0 ] ) );
                
                // Matched?
                if( $email === $address )
                {
                    // Unset using $k
                    unset( $this->{$kind}[ $k ] );
                    
                    // Don't forget to remove from ReplyToQueue
                    if( $kind === 'ReplyTo' )
                    {
                        unset( $this->ReplyToQueue[ $k ] );
                    }
                    else
                    {
                        // Clear address from all_recipients
                        unset( $this->all_recipients[ $email ] );
                        
                        // RecipientsQueue has a special structure in which a sub-array stores $kind
                        // Clear it out by using array_filter()
                        $this->RecipientsQueue = array_filter(
                            $this->RecipientsQueue,
                            function( $v ) use( $kind, $email ){
                                return ( $v[ 0 ] === $kind && strtolower( $v[ 2 ] ) === $email ? false : true );
                            }
                        );
                    }
                    
                    $return = true;
                }
            }
            
            // Not sure if important to PHPMailer so just re-index these arrays
            // Leave ReplyTo alone because it is associative, not indexed
            if(
                $return === true &&
                in_array( $kind, [ 'to', 'cc', 'bcc' ], true )
            )
            {
                $this->{$kind} = array_values( $this->{$kind} );
            }
        }
        
        return $return;
    }
    
    /**
     * Resolve $kind to an key storage position for addresses.
     *
     * @param string $kind
     *
     * @return string
     */
    public function resolveKindToKey( $kind )
    {
        // Enforce string
        // Cleanse whitespace, hyphens, and underscores
        // This allows for valid variance such as Reply-To
        // Also allows for "-B-c-C " to become "bcc" but oh well...
        $kind = ( is_scalar( $kind ) ? strtolower( preg_replace( '/[\s_-]+/', '', $kind ) ) : '' );
        
        // Acceptable $kind
        $kind = ( in_array( $kind, [ 'to', 'cc', 'bcc', 'replyto' ], true ) ? $kind : '' );
        
        // PHPMailer->ReplyTo is the correct key
        $kind = ( $kind === 'replyto' ? 'ReplyTo' : $kind );
        
        return $kind;
    }
    
    /**
     * Addresses and names can be passed in as strings, RFC822-compatible strings,
     * indexed arrays, and associative arrays.
     * Try to make sense of what was given and return an email=>name key/value pair.
     *
     * @param string|array $addresses
     * @param string|array $names
     *
     * @return bool
     */
    public function coalesceAddressesAndNames( $addresses, $names = [] )
    {
        $return = [];
        
        // explode() $addresses and $names if they're explode-able
        $addresses = ( is_scalar( $addresses ) ? explode( ',', $addresses ) : $addresses );
        $names = ( is_scalar( $names ) ? explode( ',', $names ) : $names );
        
        // If they're not iterable then make them iterable
        $addresses = ( is_array( $addresses ) ? $addresses : [ $addresses ] );
        $names = ( is_array( $names ) ? $names : [ $names ] );
        
        foreach( $addresses as $k => $v )
        {
            $address = '';
            $name = '';
            
            // Assume the key is the email if array is associative
            $address = ( !is_int( $k ) ? $k : $v );
            
            // If $addresses is associative array then assume $v contains the name
            // Otherwise, get name from $names array at this address' key position
            $name = ( !is_int( $k ) ? $v : ( isset( $names[ $k ] ) ? $names[ $k ] : $name ) );
            // Account for receiving output from PHPMailer->getReplyToAddresses() method
            if( is_array( $name ) )
            {
                $name = array_values( $name );
                $name = ( $name === '' && isset( $name[ 1 ] ) && is_scalar( $name[ 1 ] ) ? $name[ 1 ] : '' );
            }
            $name = trim( $name );
            
            // Account for receiving output from PHPMailer->getTo/Cc/Bcc/ReplyToAddresses() methods
            if( is_array( $address ) )
            {
                $address = array_values( $address );
                $name = ( $name === '' && isset( $address[ 1 ] ) && is_scalar( $address[ 1 ] ) ? $address[ 1 ] : $name );
                $address = ( isset( $address[ 0 ] ) && is_scalar( $address[ 0 ] ) ? $address[ 0 ] : '' );
            }
            
            // Account for RFC822 style address possibility
            $parseAddresses = $this->parseAddresses( $address, true, $this->CharSet );
            if( $parseAddresses )
            {
                $address = $parseAddresses[ 0 ][ 'address' ];
                $name = ( $name === '' ? $parseAddresses[ 0 ][ 'name' ] : $name );
            }
            
            $address = trim( $this->appendDefaultAddressDomain( $address ) );
            $name = trim( $name );
            
            $return[ $address ] = $name;
        }
        
        return $return;
    }
}
