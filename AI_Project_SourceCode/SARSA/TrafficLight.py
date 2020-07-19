class TrafficStateTuple:
    '''Initializing the state of a traffic Light'''
    laneLength=10
    def __init__(self,queue1,queue2,phaseTime1,phaseTime2):
        self.queue1=queue1
        self.queue2=queue2
        self.phaseTime1=phaseTime1
        self.phaseTime2=phaseTime2


